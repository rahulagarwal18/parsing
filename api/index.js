import express from "express";
import cors from "cors";
import multer from "multer";
import path from "path";
import fs from "fs";
import os from "os";
import dotenv from "dotenv";
import XLSX from "xlsx";
import { parsePdf } from "../services/llama.js";
import { 
  extractFieldsFromText, 
  convertPdfMarkdownToTable,
  detectSchemaFromMarkdown,
  generateAutoMapping,
  generateReconciliationExecutiveSummary
} from "../services/gemini.js";
import { compareData, reconcileTables } from "../services/comparator.js";
import { generateExcelReport, generatePdfReport, generateDocxReport } from "../services/exporter.js";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
dotenv.config({ path: path.resolve(__dirname, "../.env"), override: true });

const app = express();

// Enable CORS and JSON body parsing
app.use(cors());
app.use(express.json());

// Serve static frontend files from 'public' folder (primarily for local development)
app.use(express.static(path.join(process.cwd(), "public")));

// Ensure uploads and temp directories exist (use os.tmpdir() on Vercel)
const isVercel = process.env.VERCEL || process.env.NOW_BUILDER;
const uploadsDir = isVercel ? path.join(os.tmpdir(), "uploads") : "./uploads";
const tempDir = isVercel ? path.join(os.tmpdir(), "temp") : "./temp";

[uploadsDir, tempDir].forEach(dir => {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
});

// Configure Multer for file storage
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadsDir);
  },
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + "-" + Math.round(Math.random() * 1e9);
    cb(null, uniqueSuffix + "-" + file.originalname);
  }
});

const upload = multer({
  storage: storage,
  limits: {
    fileSize: 50 * 1024 * 1024 // 50MB limit per upload batch
  }
});

// In-memory data store for current job
let currentJobState = {
  excelData: null,
  excelFilePath: null,
  comparisonResults: null
};

// SSE connections map
const sseClients = new Map();

/**
 * Endpoint for SSE progress stream.
 */
app.get("/api/progress-stream", (req, res) => {
  const clientId = req.query.clientId;
  if (!clientId) {
    res.status(400).send("clientId query parameter required");
    return;
  }

  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache");
  res.setHeader("Connection", "keep-alive");
  res.flushHeaders();

  sseClients.set(clientId, res);

  // Send initial connected ping
  res.write(`data: ${JSON.stringify({ status: "connected", message: "SSE connected" })}\n\n`);

  req.on("close", () => {
    sseClients.delete(clientId);
  });
});

/**
 * Sends a message/log to a specific SSE client.
 */
function sendProgressUpdate(clientId, data) {
  const client = sseClients.get(clientId);
  if (client) {
    client.write(`data: ${JSON.stringify(data)}\n\n`);
  }
}

/**
 * Endpoint to upload Excel file.
 */
app.post("/api/upload-excel", upload.single("excelFile"), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No Excel file uploaded" });
    }

    const filePath = req.file.path;
    
    // Parse the Excel File
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // Convert to JSON row array
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
    
    // Retrieve Headers list from worksheet columns
    const headers = [];
    const range = XLSX.utils.decode_range(worksheet["!ref"]);
    const R = range.s.r; // 0-based header row
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cell_ref = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = worksheet[cell_ref];
      if (cell && cell.v !== undefined) {
        headers.push(cell.v.toString().trim());
      }
    }

    // Cache Excel data
    currentJobState.excelData = { headers, rows };
    currentJobState.excelFilePath = filePath;
    currentJobState.comparisonResults = null;

    // Send headers and first 5 rows preview to frontend
    res.json({
      success: true,
      headers,
      preview: rows.slice(0, 5)
    });
  } catch (error) {
    console.error("Excel upload error:", error);
    res.status(500).json({ error: error.message });
  }
});

/**
 * Endpoint to upload PDF files batch (up to 200+).
 * Starts background parsing and dynamic schema detection.
 */
app.post("/api/upload-pdf-batch", upload.array("pdfFiles", 205), async (req, res) => {
  const clientId = req.body.clientId;

  if (!clientId) {
    return res.status(400).json({ error: "clientId is required" });
  }
  if (!req.files || req.files.length === 0) {
    return res.status(400).json({ error: "No PDF files uploaded" });
  }

  // Respond immediately, processing continues in background
  res.json({ success: true, message: "Upload successful, processing started." });

  const files = req.files;
  const totalFiles = files.length;
  const results = [];
  const concurrencyLimit = 3;
  let detectedHeaders = null;

  sendProgressUpdate(clientId, { status: "start", total: totalFiles, message: `Starting LlamaParse extraction for ${totalFiles} files...` });

  // Process files in batches
  for (let i = 0; i < totalFiles; i += concurrencyLimit) {
    const batch = files.slice(i, i + concurrencyLimit);
    const batchPromises = batch.map(async (file, index) => {
      const fileIndex = i + index;
      const originalName = file.originalname;
      const filePath = file.path;

      try {
        sendProgressUpdate(clientId, {
          status: "parsing",
          index: fileIndex,
          filename: originalName,
          message: `[${fileIndex + 1}/${totalFiles}] Parsing PDF with LlamaParse...`
        });

        // 1. Parse PDF using LlamaParse
        const parsed = await parsePdf(filePath);

        // 2. If schema not detected yet, detect it from the first parsed PDF
        if (!detectedHeaders) {
          sendProgressUpdate(clientId, {
            status: "detecting_schema",
            index: fileIndex,
            filename: originalName,
            message: `[${fileIndex + 1}/${totalFiles}] Analyzing document structure via AI...`
          });
          detectedHeaders = await detectSchemaFromMarkdown(parsed.markdown);
          sendProgressUpdate(clientId, {
            status: "schema_detected",
            headers: detectedHeaders,
            message: `Common schema detected: ${detectedHeaders.join(", ")}`
          });
        }

        sendProgressUpdate(clientId, {
          status: "extracting",
          index: fileIndex,
          filename: originalName,
          message: `[${fileIndex + 1}/${totalFiles}] Extracting fields via Claude...`
        });

        // 3. Extract detected headers from parsed text via Claude
        const extractedData = await extractFieldsFromText(parsed.markdown, detectedHeaders);

        // Save result
        results.push({
          filename: originalName,
          rawMarkdown: parsed.markdown,
          ...extractedData
        });

        sendProgressUpdate(clientId, {
          status: "completed_file",
          index: fileIndex,
          filename: originalName,
          message: `[${fileIndex + 1}/${totalFiles}] Successfully processed file.`
        });

      } catch (err) {
        console.error(`Error processing batch PDF ${originalName}:`, err);
        sendProgressUpdate(clientId, {
          status: "failed_file",
          index: fileIndex,
          filename: originalName,
          message: `[${fileIndex + 1}/${totalFiles}] Failed: ${err.message}`
        });

        // Push empty record with filename so it shows up
        const fallbackObj = { filename: originalName, rawMarkdown: "Failed to parse PDF document: " + err.message };
        if (detectedHeaders) {
          detectedHeaders.forEach(h => fallbackObj[h] = "");
        }
        results.push(fallbackObj);
      } finally {
        try {
          fs.unlinkSync(filePath);
        } catch (unlinkErr) {
          console.error(`Failed to delete temp PDF batch file ${filePath}:`, unlinkErr);
        }
      }
    });

    await Promise.all(batchPromises);
  }

  try {
    // If we failed to parse anything or detect headers
    if (!detectedHeaders) {
      detectedHeaders = ["Invoice ID", "Date", "Amount", "Vendor"];
      results.forEach(r => {
        detectedHeaders.forEach(h => {
          if (r[h] === undefined) r[h] = "";
        });
      });
    }

    // Cache PDF parsed data
    currentJobState.pdfData = {
      headers: detectedHeaders,
      rows: results
    };
    currentJobState.comparisonResults = null; // Reset comparisons

    // Stream final results back to SSE client
    sendProgressUpdate(clientId, {
      status: "done",
      message: "Batch PDF processing completed successfully!",
      data: {
        headers: detectedHeaders,
        rows: results
      }
    });
  } catch (err) {
    console.error("Batch PDF processing error:", err);
    sendProgressUpdate(clientId, {
      status: "error",
      message: `Batch PDF processing failed: ${err.message}`
    });
  }
});

/**
 * Endpoint to auto-map PDF headers to Excel headers.
 */
app.post("/api/auto-map-columns", async (req, res) => {
  try {
    const { pdfHeaders, excelHeaders } = req.body;
    if (!pdfHeaders || !excelHeaders) {
      return res.status(400).json({ error: "pdfHeaders and excelHeaders are required" });
    }

    const mapping = await generateAutoMapping(pdfHeaders, excelHeaders);
    res.json({ success: true, mapping });
  } catch (error) {
    console.error("Auto mapping error:", error);
    res.status(500).json({ error: error.message });
  }
});

/**
 * Endpoint to reconcile PDF rows against Excel rows and generate the Claude report.
 */
app.post("/api/reconcile-tables", async (req, res) => {
  try {
    const { mappings, matchKey } = req.body;

    if (!mappings || !matchKey) {
      return res.status(400).json({ error: "mappings and matchKey are required" });
    }

    if (!currentJobState.pdfData) {
      return res.status(400).json({ error: "No PDF data found. Please parse PDFs first." });
    }

    if (!currentJobState.excelData) {
      return res.status(400).json({ error: "No Excel data found. Please upload Excel first." });
    }

    // Perform table-to-table comparison
    const reconciliationResult = reconcileTables(
      currentJobState.pdfData.rows,
      currentJobState.excelData.rows,
      mappings,
      matchKey
    );

    // Add match values to records for exporter compatibility
    reconciliationResult.records.forEach(r => {
      r.pdfMatchValue = r.pdfRow[matchKey] || "";
      const excelHeader = mappings[matchKey];
      r.excelMatchValue = r.matchedExcelRow ? r.matchedExcelRow[excelHeader] || "" : "";
    });

    // Generate Executive Report via Claude
    const executiveReport = await generateReconciliationExecutiveSummary(reconciliationResult, mappings);

    // Cache results
    currentJobState.comparisonResults = {
      report: reconciliationResult,
      selectedHeaders: Object.keys(mappings),
      matchKey: matchKey,
      mappings,
      executiveSummary: executiveReport
    };

    res.json({
      success: true,
      report: reconciliationResult,
      executiveSummary: executiveReport
    });
  } catch (error) {
    console.error("Reconciliation error:", error);
    res.status(500).json({ error: error.message });
  }
});

/**
 * Endpoint to download generated reports.
 */
app.get("/api/export/:type", async (req, res) => {
  const type = req.params.type;
  
  if (!currentJobState.comparisonResults) {
    return res.status(404).json({ error: "No comparison results available to export." });
  }

  const { report, selectedHeaders, matchKey, executiveSummary } = currentJobState.comparisonResults;

  try {
    if (type === "excel") {
      const buffer = generateExcelReport(report, selectedHeaders, matchKey);
      res.setHeader("Content-Disposition", `attachment; filename="comparison_report_${Date.now()}.xlsx"`);
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      return res.send(buffer);
    } 
    
    if (type === "pdf") {
      const buffer = await generatePdfReport(report, selectedHeaders, executiveSummary);
      res.setHeader("Content-Disposition", `attachment; filename="comparison_report_${Date.now()}.pdf"`);
      res.setHeader("Content-Type", "application/pdf");
      return res.send(buffer);
    } 
    
    if (type === "word") {
      const buffer = await generateDocxReport(report, selectedHeaders, executiveSummary);
      res.setHeader("Content-Disposition", `attachment; filename="comparison_report_${Date.now()}.docx"`);
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
      return res.send(buffer);
    }

    res.status(400).json({ error: `Unsupported export type: ${type}` });
  } catch (err) {
    console.error("Export report error:", err);
    res.status(500).json({ error: `Failed to generate report: ${err.message}` });
  }
});

export default app;

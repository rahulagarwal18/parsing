import express from "express";
import cors from "cors";
import multer from "multer";
import path from "path";
import fs from "fs";
import os from "os";
import dotenv from "dotenv";
import XLSX from "xlsx";
import { parsePdf } from "../services/llama.js";
import { extractFieldsFromText, convertPdfMarkdownToTable } from "../services/gemini.js";
import { compareData } from "../services/comparator.js";
import { generateExcelReport, generatePdfReport, generateDocxReport } from "../services/exporter.js";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
dotenv.config({ path: path.resolve(__dirname, "../.env") });

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
 * Endpoint to upload PDF as Master Document.
 */
app.post("/api/upload-pdf-master", upload.single("pdfMasterFile"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No PDF master file uploaded" });
    }

    const filePath = req.file.path;

    // 1. Parse PDF using LlamaParse
    const parsed = await parsePdf(filePath);

    // 2. Convert LlamaParse markdown text into structured table data (headers & rows) via Gemini
    const tableData = await convertPdfMarkdownToTable(parsed.markdown);

    // Clean up uploaded file
    try {
      fs.unlinkSync(filePath);
    } catch (unlinkErr) {
      console.error(`Failed to delete master PDF temp file ${filePath}:`, unlinkErr);
    }

    // Cache the structured master data in currentJobState (acting as the master sheet)
    currentJobState.excelData = {
      headers: tableData.headers,
      rows: tableData.rows
    };
    currentJobState.excelFilePath = null; // No Excel path since we uploaded a PDF
    currentJobState.comparisonResults = null;

    // Send headers and preview rows back to frontend
    res.json({
      success: true,
      headers: tableData.headers,
      preview: tableData.rows.slice(0, 5)
    });
  } catch (error) {
    console.error("PDF master upload error:", error);
    res.status(500).json({ error: error.message });
  }
});

/**
 * Endpoint to upload PDFs and trigger matching process.
 */
app.post("/api/process-pdfs", upload.array("pdfFiles", 205), async (req, res) => {
  const clientId = req.body.clientId;
  const matchKey = req.body.matchKey;
  const customPrompt = req.body.customPrompt || "";
  
  let selectedHeaders = [];
  try {
    selectedHeaders = JSON.parse(req.body.selectedHeaders || "[]");
  } catch (e) {
    return res.status(400).json({ error: "Invalid selectedHeaders format" });
  }

  if (!clientId) {
    return res.status(400).json({ error: "clientId is required" });
  }
  if (!currentJobState.excelData) {
    return res.status(400).json({ error: "No Excel data found. Please upload Excel first." });
  }
  if (!req.files || req.files.length === 0) {
    return res.status(400).json({ error: "No PDF files uploaded" });
  }
  if (!matchKey || !selectedHeaders.includes(matchKey)) {
    return res.status(400).json({ error: "Match key must be selected in comparison headers" });
  }

  // Respond immediately to the request, processing continues in background
  res.json({ success: true, message: "Processing started" });

  const files = req.files;
  const totalFiles = files.length;
  const results = [];
  const concurrencyLimit = 3;

  sendProgressUpdate(clientId, { status: "start", total: totalFiles, message: `Starting comparison pipeline for ${totalFiles} files...` });

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

        sendProgressUpdate(clientId, {
          status: "extracting",
          index: fileIndex,
          filename: originalName,
          message: `[${fileIndex + 1}/${totalFiles}] Extracting comparison fields via Gemini...`
        });

        // 2. Extract selected headers from parsed text via Gemini (with custom instructions)
        const extractedData = await extractFieldsFromText(parsed.markdown, selectedHeaders, customPrompt);

        // Save result
        results.push({
          filename: originalName,
          extractedData: extractedData
        });

        sendProgressUpdate(clientId, {
          status: "completed_file",
          index: fileIndex,
          filename: originalName,
          message: `[${fileIndex + 1}/${totalFiles}] Successfully processed file.`
        });

      } catch (err) {
        console.error(`Error processing file ${originalName}:`, err);
        sendProgressUpdate(clientId, {
          status: "failed_file",
          index: fileIndex,
          filename: originalName,
          message: `[${fileIndex + 1}/${totalFiles}] Failed: ${err.message}`
        });

        results.push({
          filename: originalName,
          extractedData: {},
          error: err.message
        });
      } finally {
        try {
          fs.unlinkSync(filePath);
        } catch (unlinkErr) {
          console.error(`Failed to delete temp file ${filePath}:`, unlinkErr);
        }
      }
    });

    await Promise.all(batchPromises);
  }

  // Once all files processed, run comparisons
  try {
    sendProgressUpdate(clientId, { status: "comparing", message: "Reconciling PDF data with Excel sheet..." });
    const finalReport = compareData(
      currentJobState.excelData.rows,
      results,
      selectedHeaders,
      matchKey
    );

    // Cache final report in server state
    currentJobState.comparisonResults = {
      report: finalReport,
      selectedHeaders,
      matchKey
    };

    // Send final results back to SSE client
    sendProgressUpdate(clientId, {
      status: "done",
      message: "Processing completed successfully!",
      data: finalReport
    });
  } catch (err) {
    console.error("Comparison reconciliation error:", err);
    sendProgressUpdate(clientId, {
      status: "error",
      message: `Reconciliation failed: ${err.message}`
    });
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

  const { report, selectedHeaders, matchKey } = currentJobState.comparisonResults;

  try {
    if (type === "excel") {
      const buffer = generateExcelReport(report, selectedHeaders, matchKey);
      res.setHeader("Content-Disposition", `attachment; filename="comparison_report_${Date.now()}.xlsx"`);
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      return res.send(buffer);
    } 
    
    if (type === "pdf") {
      const buffer = await generatePdfReport(report, selectedHeaders);
      res.setHeader("Content-Disposition", `attachment; filename="comparison_report_${Date.now()}.pdf"`);
      res.setHeader("Content-Type", "application/pdf");
      return res.send(buffer);
    } 
    
    if (type === "word") {
      const buffer = await generateDocxReport(report, selectedHeaders);
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

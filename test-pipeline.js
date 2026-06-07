import { reconcileTables } from "./services/comparator.js";
import { generateExcelReport, generatePdfReport, generateDocxReport } from "./services/exporter.js";
import fs from "fs";

async function runTest() {
  console.log("=== STARTING NEW RECONCILIATION PIPELINE VALIDATION TESTS ===");
  
  if (!fs.existsSync("./temp")) {
    fs.mkdirSync("./temp", { recursive: true });
  }

  // 1. Mock Excel Rows
  const excelRows = [
    { "Invoice ID": "INV-001", "Vendor": "Acme Corp", "Amount": "1,250.00", "Date": "2026-06-01" },
    { "Invoice ID": "INV-002", "Vendor": "Globex Corp", "Amount": "$850.50", "Date": "2026-06-03" },
    { "Invoice ID": "INV-003", "Vendor": "Umbrella Corp", "Amount": "5,000", "Date": "2026-06-05" },
    { "Invoice ID": "INV-004", "Vendor": "Initech Corp", "Amount": "320.00", "Date": "2026-06-06" }
  ];

  // 2. Mock Parsed PDF Rows
  const pdfRows = [
    { filename: "invoice_acme.pdf", "Invoice ID": "INV-001", "Vendor": "Acme Corp", "Amount": "$1250", "Date": "2026-06-01" }, // Full Match
    { filename: "invoice_globex.pdf", "Invoice ID": "INV-002", "Vendor": "Globex Corporation", "Amount": "850.50", "Date": "2026-06-03" }, // Mismatch on Vendor
    { filename: "invoice_umbrella.pdf", "Invoice ID": "INV-003", "Vendor": "Umbrella Corp", "Amount": "", "Date": "2026-06-05" }, // Partial match (Missing Amount)
    { filename: "invoice_unknown.pdf", "Invoice ID": "INV-999", "Vendor": "Stark Industries", "Amount": "9999", "Date": "2026-06-07" } // No Excel Row Match
  ];

  const mappings = {
    "Invoice ID": "Invoice ID",
    "Vendor": "Vendor",
    "Amount": "Amount",
    "Date": "Date"
  };
  const matchKey = "Invoice ID";

  console.log("Running reconciliation engine comparison...");
  const report = reconcileTables(pdfRows, excelRows, mappings, matchKey);

  // Add match values to records for exporter compatibility
  report.records.forEach(r => {
    r.pdfMatchValue = r.pdfRow[matchKey] || "";
    r.excelMatchValue = r.matchedExcelRow ? r.matchedExcelRow[mappings[matchKey]] || "" : "";
  });

  // Validate Summary Output
  console.log("Summary Output:", report.summary);
  if (report.summary.totalPdfsProcessed !== 4) throw new Error("Expected 4 processed records");
  if (report.summary.totalMatches !== 2) throw new Error("Expected 2 matches (1 full, 1 partial)");
  if (report.summary.totalMismatches !== 1) throw new Error("Expected 1 mismatch");
  if (report.summary.totalNoExcelMatch !== 1) throw new Error("Expected 1 unmatched PDF row");
  if (report.summary.totalUnmatchedExcelRows !== 1) throw new Error("Expected 1 unmatched Excel row");

  console.log("Reconciliation table-to-table comparison matches expected outcomes.");

  const selectedHeaders = Object.keys(mappings);

  // Test Report Exporters
  console.log("Testing Excel Exporter...");
  const excelBuffer = generateExcelReport(report, selectedHeaders, matchKey);
  fs.writeFileSync("./temp/test_report.xlsx", excelBuffer);
  console.log(" Excel report generated: ./temp/test_report.xlsx");

  const mockExecutiveSummary = "# Audit Summary\n\nThis is a mock executive reconciliation summary.\n\n## Findings\n- Acme Corp matched successfully.\n- Globex Corp has a vendor mismatch.\n- Umbrella Corp is missing its amount.\n\n### Recommendation\nCheck Globex Corp vendor name.";

  console.log("Testing PDF Exporter...");
  const pdfBuffer = await generatePdfReport(report, selectedHeaders, mockExecutiveSummary);
  fs.writeFileSync("./temp/test_report.pdf", pdfBuffer);
  console.log(" PDF report generated: ./temp/test_report.pdf");

  console.log("Testing Word Exporter...");
  const docxBuffer = await generateDocxReport(report, selectedHeaders, mockExecutiveSummary);
  fs.writeFileSync("./temp/test_report.docx", docxBuffer);
  console.log(" Word report generated: ./temp/test_report.docx");

  console.log("=== ALL PIPELINE VALIDATION TESTS PASSED SUCCESSFULLY ===");
}

runTest().catch(err => {
  console.error("Test failed:", err);
  process.exit(1);
});

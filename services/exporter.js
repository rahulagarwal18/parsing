import XLSX from "xlsx";
import PDFDocument from "pdfkit";
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType } from "docx";

/**
 * Generates an Excel report using the 'xlsx' library.
 * 
 * @param {Object} comparisonResult - Result from comparator.reconcileTables
 * @param {string[]} selectedHeaders - List of headers compared
 * @param {string} matchKey - Match key
 * @returns {Buffer} Excel file buffer
 */
export function generateExcelReport(comparisonResult, selectedHeaders, matchKey) {
  const wb = XLSX.utils.book_new();

  // --- SHEET 1: SUMMARY ---
  const summaryData = [
    ["PDF & Excel Data Comparison Report - Summary"],
    [],
    ["Metric", "Value"],
    ["Total PDFs Processed", comparisonResult.summary.totalPdfsProcessed],
    ["Full & Partial Matches", comparisonResult.summary.totalMatches],
    ["Mismatches Found", comparisonResult.summary.totalMismatches],
    ["PDFs with No Excel Row Match", comparisonResult.summary.totalNoExcelMatch],
    ["Unmatched Rows in Excel", comparisonResult.summary.totalUnmatchedExcelRows],
    [],
    ["Report Generated At", new Date().toLocaleString()]
  ];
  const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
  XLSX.utils.book_append_sheet(wb, summarySheet, "Summary");

  // --- SHEET 2: COMPARISON DETAILS ---
  const detailHeaders = ["PDF Filename", "Status", "Matched Excel Row", `Key Value (${matchKey})`];
  for (const h of selectedHeaders) {
    detailHeaders.push(`Excel: ${h}`, `PDF: ${h}`, `Status: ${h}`);
  }

  const detailRows = [detailHeaders];

  for (const record of comparisonResult.records) {
    const keyValue = record.pdfMatchValue || (record.matchedExcelRow ? record.matchedExcelRow[matchKey] || "N/A" : "N/A");
    const row = [
      record.filename,
      record.status,
      record.excelRowIndex || "N/A",
      keyValue
    ];

    for (const h of selectedHeaders) {
      const field = record.fields[h] || {};
      row.push(
        field.excelValue !== undefined ? field.excelValue : "",
        field.pdfValue !== undefined ? field.pdfValue : "",
        field.status || "N/A"
      );
    }
    detailRows.push(row);
  }

  const detailSheet = XLSX.utils.aoa_to_sheet(detailRows);
  XLSX.utils.book_append_sheet(wb, detailSheet, "Comparison Details");

  // --- SHEET 3: UNMATCHED EXCEL ROWS ---
  const unmatchedRows = [["Original Excel Row Number"]];
  if (comparisonResult.unmatchedExcelRows && comparisonResult.unmatchedExcelRows.length > 0) {
    const sampleRow = comparisonResult.unmatchedExcelRows[0].data;
    const excelHeaders = Object.keys(sampleRow);
    unmatchedRows[0].push(...excelHeaders);

    for (const item of comparisonResult.unmatchedExcelRows) {
      const row = [item.excelRowIndex];
      for (const h of excelHeaders) {
        row.push(item.data[h] !== undefined ? item.data[h] : "");
      }
      unmatchedRows.push(row);
    }
  } else {
    unmatchedRows.push(["No unmatched rows. All Excel records matched a PDF!"]);
  }

  const unmatchedSheet = XLSX.utils.aoa_to_sheet(unmatchedRows);
  XLSX.utils.book_append_sheet(wb, unmatchedSheet, "Unmatched Excel Rows");

  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
}

/**
 * Generates a PDF report using 'pdfkit'.
 * 
 * @param {Object} comparisonResult 
 * @param {string[]} selectedHeaders 
 * @param {string} executiveSummary
 * @returns {Promise<Buffer>} PDF file buffer
 */
export function generatePdfReport(comparisonResult, selectedHeaders, executiveSummary) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ margin: 40, size: "A4" });
    const chunks = [];

    doc.on("data", chunk => chunks.push(chunk));
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", err => reject(err));

    // Colors
    const primaryColor = "#6366F1"; // Indigo
    const darkTextColor = "#1E293B"; // Slate 800
    const lightTextColor = "#64748B"; // Slate 500
    const matchColor = "#10B981"; // Emerald
    const mismatchColor = "#EF4444"; // Red
    const warningColor = "#F59E0B"; // Amber

    // --- TITLE PAGE ---
    doc.rect(0, 0, doc.page.width, 25).fill(primaryColor);

    doc.moveDown(2);
    doc.fillColor(darkTextColor).font("Helvetica-Bold").fontSize(26).text("Data Comparison Report", { align: "center" });
    doc.fillColor(lightTextColor).font("Helvetica").fontSize(12).text("Automated PDF vs Excel Reconciliation", { align: "center" });
    doc.moveDown(1);
    doc.text(`Generated on: ${new Date().toLocaleString()}`, { align: "center" });
    
    doc.moveDown(2);
    doc.rect(40, doc.y, doc.page.width - 80, 1).fill("#E2E8F0");
    doc.moveDown(2);

    // Summary Statistics Cards
    doc.fillColor(darkTextColor).font("Helvetica-Bold").fontSize(16).text("Reconciliation Summary");
    doc.moveDown(0.5);

    const startY = doc.y;
    const cardWidth = (doc.page.width - 100) / 3;

    // Card 1: Processed
    doc.rect(40, startY, cardWidth, 70).fill("#F8FAFC");
    doc.fillColor(darkTextColor).font("Helvetica-Bold").fontSize(20).text(String(comparisonResult.summary.totalPdfsProcessed), 55, startY + 15);
    doc.fillColor(lightTextColor).font("Helvetica").fontSize(10).text("PDFs Processed", 55, startY + 45);

    // Card 2: Matches
    doc.rect(40 + cardWidth + 10, startY, cardWidth, 70).fill("#F0FDF4");
    doc.fillColor(matchColor).font("Helvetica-Bold").fontSize(20).text(String(comparisonResult.summary.totalMatches), 40 + cardWidth + 25, startY + 15);
    doc.fillColor(lightTextColor).font("Helvetica").fontSize(10).text("Matches / Partial", 40 + cardWidth + 25, startY + 45);

    // Card 3: Mismatches
    doc.rect(40 + 2 * (cardWidth + 10), startY, cardWidth, 70).fill("#FEF2F2");
    doc.fillColor(mismatchColor).font("Helvetica-Bold").fontSize(20).text(String(comparisonResult.summary.totalMismatches), 40 + 2 * (cardWidth + 10) + 15, startY + 15);
    doc.fillColor(lightTextColor).font("Helvetica").fontSize(10).text("Mismatches Found", 40 + 2 * (cardWidth + 10) + 15, startY + 45);

    doc.y = startY + 90;

    // Card 4: Unmatched Excel
    const bottomCardWidth = (doc.page.width - 90) / 2;
    const bottomY = doc.y;
    doc.rect(40, bottomY, bottomCardWidth, 60).fill("#F8FAFC");
    doc.fillColor(darkTextColor).font("Helvetica-Bold").fontSize(16).text(String(comparisonResult.summary.totalNoExcelMatch), 55, bottomY + 12);
    doc.fillColor(lightTextColor).font("Helvetica").fontSize(9).text("PDFs with no matching Excel Row", 55, bottomY + 35);

    // Card 5: Unmatched Rows
    doc.rect(40 + bottomCardWidth + 10, bottomY, bottomCardWidth, 60).fill("#FFFBEB");
    doc.fillColor(warningColor).font("Helvetica-Bold").fontSize(16).text(String(comparisonResult.summary.totalUnmatchedExcelRows), 40 + bottomCardWidth + 25, bottomY + 12);
    doc.fillColor(lightTextColor).font("Helvetica").fontSize(9).text("Unmatched rows remaining in Excel", 40 + bottomCardWidth + 25, bottomY + 35);

    doc.y = bottomY + 90;

    // --- EXECUTIVE SUMMARY SECTION ---
    if (executiveSummary) {
      doc.addPage();
      doc.fillColor(darkTextColor).font("Helvetica-Bold").fontSize(18).text("Executive Reconciliation Summary");
      doc.moveDown(1);
      
      const cleanSummary = executiveSummary
        .replace(/```markdown\s*/g, "")
        .replace(/```\s*/g, "")
        .trim();

      const lines = cleanSummary.split("\n");
      for (const line of lines) {
        const trimmed = line.trim();
        if (trimmed.startsWith("###")) {
          doc.moveDown(0.5);
          doc.fillColor(primaryColor).font("Helvetica-Bold").fontSize(12).text(trimmed.replace(/^###\s+/, ""), { width: doc.page.width - 80 });
          doc.moveDown(0.3);
        } else if (trimmed.startsWith("##")) {
          doc.moveDown(0.8);
          doc.fillColor(primaryColor).font("Helvetica-Bold").fontSize(14).text(trimmed.replace(/^##\s+/, ""), { width: doc.page.width - 80 });
          doc.moveDown(0.4);
        } else if (trimmed.startsWith("#")) {
          doc.moveDown(1);
          doc.fillColor(primaryColor).font("Helvetica-Bold").fontSize(16).text(trimmed.replace(/^#\s+/, ""), { width: doc.page.width - 80 });
          doc.moveDown(0.5);
        } else if (trimmed.startsWith("-") || trimmed.startsWith("*")) {
          doc.fillColor(darkTextColor).font("Helvetica").fontSize(10);
          doc.text("• " + trimmed.substring(1).trim(), { width: doc.page.width - 80, paragraphGap: 4 });
        } else if (trimmed.length > 0) {
          doc.fillColor(darkTextColor).font("Helvetica").fontSize(10);
          doc.text(trimmed, { width: doc.page.width - 80, paragraphGap: 6 });
        } else {
          doc.moveDown(0.3);
        }
      }
    }

    // --- DETAILED RECORDS TABLE ---
    doc.addPage();
    doc.fillColor(darkTextColor).font("Helvetica-Bold").fontSize(16).text("Detailed Comparison Results");
    doc.moveDown(0.5);

    // Draw table headers
    const colWidths = [180, 80, 100, 150]; // Filename, Row, Status, Issue Details
    const headers = ["PDF Filename", "Excel Row", "Status", "Comparison Notes"];
    let curX = 40;
    
    // Draw table header background
    const headerY = doc.y + 5;
    doc.rect(40, doc.y, doc.page.width - 80, 20).fill(primaryColor);
    doc.fillColor("#FFFFFF").font("Helvetica-Bold").fontSize(9);

    for (let i = 0; i < headers.length; i++) {
      doc.text(headers[i], curX + 5, headerY, { width: colWidths[i] - 10, lineBreak: false });
      curX += colWidths[i];
    }
    
    doc.y += 20;
    doc.font("Helvetica").fontSize(8).fillColor(darkTextColor);

    // Draw rows
    for (const record of comparisonResult.records) {
      // Check if page overflow
      if (doc.y > doc.page.height - 60) {
        doc.addPage();
        // Redraw table header on new page
        const newHeaderY = doc.y + 5;
        doc.rect(40, doc.y, doc.page.width - 80, 20).fill(primaryColor);
        doc.fillColor("#FFFFFF").font("Helvetica-Bold").fontSize(9);
        let tempX = 40;
        for (let i = 0; i < headers.length; i++) {
          doc.text(headers[i], tempX + 5, newHeaderY, { width: colWidths[i] - 10, lineBreak: false });
          tempX += colWidths[i];
        }
        doc.y += 20;
        doc.font("Helvetica").fontSize(8).fillColor(darkTextColor);
      }

      const rowY = doc.y;
      
      // Calculate mismatch details
      const issues = [];
      if (record.status === "MISMATCH") {
        for (const h of selectedHeaders) {
          const field = record.fields[h];
          if (field && field.status === "MISMATCH") {
            issues.push(`${h}: Excel '${field.excelValue}' vs PDF '${field.pdfValue}'`);
          }
        }
      } else if (record.status === "PARTIAL_MATCH") {
        for (const h of selectedHeaders) {
          const field = record.fields[h];
          if (field && field.status === "MISSING_IN_PDF") {
            issues.push(`${h} is missing in PDF`);
          }
        }
      } else if (record.status === "NO_EXCEL_MATCH") {
        issues.push("Could not find matching row key in Excel sheet.");
      } else {
        issues.push("All compared fields match perfectly.");
      }

      const issuesText = issues.join("; ");
      
      // Status Color
      let statusColor = darkTextColor;
      if (record.status === "FULL_MATCH") statusColor = matchColor;
      else if (record.status === "MISMATCH") statusColor = mismatchColor;
      else if (record.status === "PARTIAL_MATCH" || record.status === "NO_EXCEL_MATCH") statusColor = warningColor;

      // Draw Row Border
      doc.rect(40, rowY, doc.page.width - 80, 30).stroke("#F1F5F9");

      // Write values
      doc.fillColor(darkTextColor).font("Helvetica-Bold").text(record.filename, 45, rowY + 10, { width: colWidths[0] - 10, height: 18, ellipsis: true });
      doc.font("Helvetica").text(record.excelRowIndex ? `Row ${record.excelRowIndex}` : "N/A", 40 + colWidths[0] + 5, rowY + 10);
      
      doc.fillColor(statusColor).font("Helvetica-Bold").text(record.status, 40 + colWidths[0] + colWidths[1] + 5, rowY + 10);
      
      doc.fillColor(lightTextColor).font("Helvetica").text(issuesText, 40 + colWidths[0] + colWidths[1] + colWidths[2] + 5, rowY + 6, { width: colWidths[3] - 10, height: 22, ellipsis: true });

      doc.y = rowY + 30;
    }

    doc.end();
  });
}

/**
 * Generates a Word report using the 'docx' library.
 * 
 * @param {Object} comparisonResult 
 * @param {string[]} selectedHeaders 
 * @param {string} executiveSummary
 * @returns {Promise<Buffer>} Word document buffer
 */
export async function generateDocxReport(comparisonResult, selectedHeaders, executiveSummary) {
  const tableRows = [
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "PDF Filename", bold: true, color: "FFFFFF" })] })], shading: { fill: "6366F1" } }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Excel Row", bold: true, color: "FFFFFF" })] })], shading: { fill: "6366F1" } }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Status", bold: true, color: "FFFFFF" })] })], shading: { fill: "6366F1" } }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Comparison Notes", bold: true, color: "FFFFFF" })] })], shading: { fill: "6366F1" } }),
      ]
    })
  ];

  for (const record of comparisonResult.records) {
    const issues = [];
    if (record.status === "MISMATCH") {
      for (const h of selectedHeaders) {
        const field = record.fields[h];
        if (field && field.status === "MISMATCH") {
          issues.push(`${h}: Excel [${field.excelValue}] != PDF [${field.pdfValue}]`);
        }
      }
    } else if (record.status === "PARTIAL_MATCH") {
      for (const h of selectedHeaders) {
        const field = record.fields[h];
        if (field && field.status === "MISSING_IN_PDF") {
          issues.push(`${h} is missing in PDF`);
        }
      }
    } else if (record.status === "NO_EXCEL_MATCH") {
      issues.push("No Excel match found.");
    } else {
      issues.push("Perfect match.");
    }

    tableRows.push(
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph(record.filename)] }),
          new TableCell({ children: [new Paragraph(record.excelRowIndex ? `Row ${record.excelRowIndex}` : "N/A")] }),
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: record.status, bold: true })] })] }),
          new TableCell({ children: [new Paragraph(issues.join(", "))] }),
        ]
      })
    );
  }

  const docChildren = [
    new Paragraph({
      children: [
        new TextRun({
          text: "PDF & Excel Reconciliation Report",
          bold: true,
          size: 36,
          color: "6366F1"
        })
      ]
    }),
    new Paragraph({ text: `Report Generated At: ${new Date().toLocaleString()}` }),
    new Paragraph({ text: "" }),
    
    new Paragraph({
      children: [new TextRun({ text: "Summary Metrics", bold: true, size: 24 })]
    }),
    new Paragraph({ text: `• Total PDFs Processed: ${comparisonResult.summary.totalPdfsProcessed}` }),
    new Paragraph({ text: `• Full & Partial Matches: ${comparisonResult.summary.totalMatches}` }),
    new Paragraph({ text: `• Mismatches: ${comparisonResult.summary.totalMismatches}` }),
    new Paragraph({ text: `• PDFs with No Excel Match: ${comparisonResult.summary.totalNoExcelMatch}` }),
    new Paragraph({ text: `• Unmatched Excel Rows: ${comparisonResult.summary.totalUnmatchedExcelRows}` }),
    new Paragraph({ text: "" }),
  ];

  if (executiveSummary) {
    docChildren.push(
      new Paragraph({
        children: [new TextRun({ text: "Executive Reconciliation Summary", bold: true, size: 24, color: "6366F1" })]
      }),
      new Paragraph({ text: "" })
    );

    const cleanSummary = executiveSummary
      .replace(/```markdown\s*/g, "")
      .replace(/```\s*/g, "")
      .trim();

    const lines = cleanSummary.split("\n");
    lines.forEach(line => {
      const trimmed = line.trim();
      if (trimmed.startsWith("###")) {
        docChildren.push(new Paragraph({ children: [new TextRun({ text: trimmed.replace(/^###\s+/, ""), bold: true, size: 20, color: "6366F1" })], spacing: { before: 200, after: 100 } }));
      } else if (trimmed.startsWith("##")) {
        docChildren.push(new Paragraph({ children: [new TextRun({ text: trimmed.replace(/^##\s+/, ""), bold: true, size: 24, color: "6366F1" })], spacing: { before: 240, after: 120 } }));
      } else if (trimmed.startsWith("#")) {
        docChildren.push(new Paragraph({ children: [new TextRun({ text: trimmed.replace(/^#\s+/, ""), bold: true, size: 28, color: "6366F1" })], spacing: { before: 300, after: 150 } }));
      } else if (trimmed.startsWith("-") || trimmed.startsWith("*")) {
        docChildren.push(new Paragraph({ text: `• ${trimmed.substring(1).trim()}`, spacing: { after: 100 } }));
      } else if (trimmed.length > 0) {
        docChildren.push(new Paragraph({ text: trimmed, spacing: { after: 120 } }));
      }
    });
    docChildren.push(new Paragraph({ text: "" }));
  }

  docChildren.push(
    new Paragraph({
      children: [new TextRun({ text: "Comparison Log Table", bold: true, size: 24 })]
    }),
    new Paragraph({ text: "" }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: tableRows
    })
  );

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: docChildren
      }
    ]
  });

  return await Packer.toBuffer(doc);
}

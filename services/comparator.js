/**
 * Normalizes text values to make comparisons robust (trims, handles casing, standardizes spacing).
 * @param {any} val
 * @returns {string}
 */
export function normalizeValue(val) {
  if (val === undefined || val === null) return "";
  return String(val)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " "); // collapse multiple spaces
}

/**
 * Normalizes numeric values (removes currency symbols, commas, spaces and compares as numbers if possible).
 * @param {string} val
 * @returns {number | string}
 */
export function normalizeNumeric(val) {
  const clean = String(val)
    .replace(/[\$,€,£]/g, "") // remove currency symbols
    .replace(/,/g, "")        // remove thousand separators
    .trim();
  const num = parseFloat(clean);
  return isNaN(num) ? normalizeValue(val) : num;
}

/**
 * Compares two values for a match, handling numeric variations.
 * @param {any} excelVal 
 * @param {any} pdfVal 
 * @returns {boolean}
 */
export function valuesMatch(excelVal, pdfVal) {
  const normExcel = normalizeValue(excelVal);
  const normPdf = normalizeValue(pdfVal);

  if (normExcel === normPdf) return true;

  // Try numeric comparison
  const numExcel = normalizeNumeric(excelVal);
  const numPdf = normalizeNumeric(pdfVal);
  if (typeof numExcel === "number" && typeof numPdf === "number") {
    return numExcel === numPdf;
  }

  return false;
}

/**
 * Compares PDF extracted data with Excel rows.
 * 
 * @param {Object[]} excelRows - The rows parsed from Excel (array of objects)
 * @param {Object[]} pdfResults - Array of { filename, extractedData } where extractedData is a key-value map
 * @param {string[]} selectedHeaders - Headers to compare
 * @param {string} matchKey - The header designated as the unique identifier/match key
 * @returns {Object} Comparison report summary and detailed records
 */
export function compareData(excelRows, pdfResults, selectedHeaders, matchKey) {
  const records = [];
  let totalMatches = 0;
  let totalMismatches = 0;
  let totalNoExcelMatch = 0;

  for (const pdf of pdfResults) {
    const filename = pdf.filename;
    const pdfData = pdf.extractedData || {};
    const pdfMatchVal = pdfData[matchKey];

    // Find the matching row in Excel
    let matchedRow = null;
    let matchedRowIndex = -1;

    if (pdfMatchVal) {
      matchedRowIndex = excelRows.findIndex(row => {
        return valuesMatch(row[matchKey], pdfMatchVal);
      });
      if (matchedRowIndex !== -1) {
        matchedRow = excelRows[matchedRowIndex];
      }
    }

    const fieldComparisons = {};
    let isFullMatch = true;
    let hasMismatch = false;
    let hasMissing = false;

    if (matchedRow) {
      // Compare each selected header
      for (const header of selectedHeaders) {
        const excelVal = matchedRow[header] !== undefined ? matchedRow[header] : "";
        const pdfVal = pdfData[header] !== undefined ? pdfData[header] : "";

        const isMatch = valuesMatch(excelVal, pdfVal);
        const isEmptyPdf = normalizeValue(pdfVal) === "";

        let status = "MATCH";
        if (!isMatch) {
          if (isEmptyPdf) {
            status = "MISSING_IN_PDF";
            hasMissing = true;
            isFullMatch = false;
          } else {
            status = "MISMATCH";
            hasMismatch = true;
            isFullMatch = false;
          }
        }

        fieldComparisons[header] = {
          excelValue: excelVal,
          pdfValue: pdfVal,
          status: status
        };
      }

      let statusSummary = "FULL_MATCH";
      if (hasMismatch) {
        statusSummary = "MISMATCH";
        totalMismatches++;
      } else if (hasMissing) {
        statusSummary = "PARTIAL_MATCH"; // matches what is there, but some are missing
        totalMatches++;
      } else {
        totalMatches++;
      }

      records.push({
        filename,
        excelRowIndex: matchedRowIndex + 2, // 1-based index, accounting for header row
        matchedExcelRow: matchedRow,
        status: statusSummary,
        fields: fieldComparisons
      });
    } else {
      // No match found in Excel
      totalNoExcelMatch++;
      // Fill comparison fields as missing/unmatched
      for (const header of selectedHeaders) {
        fieldComparisons[header] = {
          excelValue: "N/A (No Row Match)",
          pdfValue: pdfData[header] || "",
          status: "NO_EXCEL_MATCH"
        };
      }

      records.push({
        filename,
        excelRowIndex: null,
        matchedExcelRow: null,
        status: "NO_EXCEL_MATCH",
        fields: fieldComparisons
      });
    }
  }

  // Find Excel rows that were NOT matched by any PDF
  const matchedExcelRowIndices = new Set(
    records
      .filter(r => r.excelRowIndex !== null)
      .map(r => r.excelRowIndex - 2) // back to 0-based
  );

  const unmatchedExcelRows = [];
  excelRows.forEach((row, idx) => {
    if (!matchedExcelRowIndices.has(idx)) {
      unmatchedExcelRows.push({
        excelRowIndex: idx + 2,
        data: row
      });
    }
  });

  return {
    summary: {
      totalPdfsProcessed: pdfResults.length,
      totalMatches,
      totalMismatches,
      totalNoExcelMatch,
      totalUnmatchedExcelRows: unmatchedExcelRows.length
    },
    records,
  };
}

/**
 * Reconciles parsed PDF rows against Excel rows based on mapping definitions.
 * 
 * @param {Object[]} pdfRows - The records extracted from PDFs (array of objects)
 * @param {Object[]} excelRows - The rows parsed from Excel (array of objects)
 * @param {Object} mappings - Mapping of PDF Header -> Excel Header
 * @param {string} matchKey - PDF header designated as the unique matching key
 * @returns {Object} Comparison report containing stats, detailed records, and unmatched rows
 */
export function reconcileTables(pdfRows, excelRows, mappings, matchKey) {
  const excelMatchKey = mappings[matchKey];
  const pdfHeaders = Object.keys(mappings);
  
  const records = [];
  let totalMatches = 0;
  let totalMismatches = 0;
  let totalNoExcelMatch = 0;

  // Track Excel rows that have been matched
  const matchedExcelRowIndices = new Set();

  pdfRows.forEach((pdfRow, idx) => {
    const filename = pdfRow.filename || `Record ${idx + 1}`;
    const pdfMatchVal = pdfRow[matchKey];

    // Find matching Excel row
    let matchedRow = null;
    let matchedRowIndex = -1;

    if (pdfMatchVal !== undefined && pdfMatchVal !== null && String(pdfMatchVal).trim() !== "") {
      matchedRowIndex = excelRows.findIndex(row => {
        return valuesMatch(row[excelMatchKey], pdfMatchVal);
      });
      if (matchedRowIndex !== -1) {
        matchedRow = excelRows[matchedRowIndex];
        matchedExcelRowIndices.add(matchedRowIndex);
      }
    }

    const fieldComparisons = {};
    let isFullMatch = true;
    let hasMismatch = false;
    let hasMissing = false;

    // We only compare if we matched a row
    if (matchedRow) {
      pdfHeaders.forEach(pdfHeader => {
        const excelHeader = mappings[pdfHeader];
        // If there's no mapping for a header, skip comparison on it
        if (!excelHeader) {
          fieldComparisons[pdfHeader] = {
            excelHeader: "",
            excelValue: "N/A (Not Mapped)",
            pdfValue: pdfRow[pdfHeader] || "",
            status: "MATCH"
          };
          return;
        }

        const excelVal = matchedRow[excelHeader] !== undefined ? matchedRow[excelHeader] : "";
        const pdfVal = pdfRow[pdfHeader] !== undefined ? pdfRow[pdfHeader] : "";

        const isMatch = valuesMatch(excelVal, pdfVal);
        const isEmptyPdf = normalizeValue(pdfVal) === "";

        let status = "MATCH";
        if (!isMatch) {
          if (isEmptyPdf) {
            status = "MISSING_IN_PDF";
            hasMissing = true;
            isFullMatch = false;
          } else {
            status = "MISMATCH";
            hasMismatch = true;
            isFullMatch = false;
          }
        }

        fieldComparisons[pdfHeader] = {
          excelHeader: excelHeader,
          excelValue: excelVal,
          pdfValue: pdfVal,
          status: status
        };
      });

      let statusSummary = "FULL_MATCH";
      if (hasMismatch) {
        statusSummary = "MISMATCH";
        totalMismatches++;
      } else if (hasMissing) {
        statusSummary = "PARTIAL_MATCH";
        totalMatches++;
      } else {
        totalMatches++;
      }

      records.push({
        filename,
        pdfRowIndex: idx + 1,
        excelRowIndex: matchedRowIndex + 2, // 1-based index, accounting for header row
        matchedExcelRow: matchedRow,
        pdfRow: pdfRow,
        status: statusSummary,
        fields: fieldComparisons
      });
    } else {
      // No match found in Excel
      totalNoExcelMatch++;
      pdfHeaders.forEach(pdfHeader => {
        fieldComparisons[pdfHeader] = {
          excelHeader: mappings[pdfHeader] || "",
          excelValue: "N/A (No Row Match)",
          pdfValue: pdfRow[pdfHeader] || "",
          status: "NO_EXCEL_MATCH"
        };
      });

      records.push({
        filename,
        pdfRowIndex: idx + 1,
        excelRowIndex: null,
        matchedExcelRow: null,
        pdfRow: pdfRow,
        status: "NO_EXCEL_MATCH",
        fields: fieldComparisons
      });
    }
  });

  // Collect unmatched Excel rows
  const unmatchedExcelRows = [];
  excelRows.forEach((row, idx) => {
    if (!matchedExcelRowIndices.has(idx)) {
      unmatchedExcelRows.push({
        excelRowIndex: idx + 2,
        data: row
      });
    }
  });

  return {
    summary: {
      totalPdfsProcessed: pdfRows.length,
      totalMatches,
      totalMismatches,
      totalNoExcelMatch,
      totalUnmatchedExcelRows: unmatchedExcelRows.length
    },
    records,
    unmatchedExcelRows
  };
}


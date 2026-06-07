import dotenv from "dotenv";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
dotenv.config({ path: path.resolve(__dirname, "../.env"), override: true });

/**
 * Normalizes a JSON schema to use lowercase types for Anthropic Tool schemas.
 */
function normalizeSchema(schema) {
  return JSON.parse(JSON.stringify(schema, (key, value) => {
    if (key === "type" && typeof value === "string") {
      return value.toLowerCase();
    }
    return value;
  }));
}
/**
 * Dynamically fetches available models from Anthropic API and sorts them
 * to prioritize Sonnet and Haiku (faster/cheaper) over Opus.
 */
async function getAvailableModels(apiKey) {
  try {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 3000);

    const response = await fetch("https://api.anthropic.com/v1/models", {
      method: "GET",
      headers: {
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
      },
      signal: controller.signal
    });

    clearTimeout(timeoutId);

    if (response.ok) {
      const data = await response.json();
      if (data && Array.isArray(data.data)) {
        const activeModels = data.data.map(m => m.id);
        // Sort models so Sonnet/Haiku are prioritized over Opus/others to optimize speed & cost
        activeModels.sort((a, b) => {
          const score = (name) => {
            const nameLower = name.toLowerCase();
            if (nameLower.includes("sonnet")) return 1;
            if (nameLower.includes("haiku")) return 2;
            if (nameLower.includes("opus")) return 3;
            return 4;
          };
          return score(a) - score(b);
        });
        return activeModels;
      }
    }
  } catch (err) {
    console.warn("Failed to fetch available Anthropic models dynamically:", err.message);
  }
  return [];
}

/**
 * Calls the Anthropic Claude API using native fetch with tool choice forcing.
 * This guarantees a structured JSON output conforming to the responseSchema.
 */
async function callClaude(prompt, responseSchema, toolName = "extract_data") {
  const apiKey = process.env.ANTHROPIC_API_KEY || process.env.GEMINI_API_KEY;
  if (!apiKey) {
    throw new Error("Anthropic API key is not configured.");
  }

  const normalizedSchema = normalizeSchema(responseSchema);
  
  // Dynamically fetch models available to this key
  const fetchedModels = await getAvailableModels(apiKey);
  
  // List of fallback models to try in order of preference
  const defaultModels = [
    "claude-sonnet-4-6",
    "claude-sonnet-4-5-20250929",
    "claude-haiku-4-5-20251001",
    "claude-3-5-sonnet-latest",
    "claude-3-5-sonnet-20241022",
    "claude-3-5-sonnet-20240620",
    "claude-3-5-haiku-latest",
    "claude-3-5-haiku-20241022",
    "claude-3-haiku-20240307"
  ];

  // Merge lists to prioritize fetched active models, removing duplicates
  const models = [...new Set([...fetchedModels, ...defaultModels])];

  let lastError = null;

  for (const model of models) {
    const requestBody = {
      model: model,
      max_tokens: 4000,
      messages: [
        { role: "user", content: prompt }
      ],
      tools: [
        {
          name: toolName,
          description: "Format the extracted data into structured JSON matching the schema.",
          input_schema: normalizedSchema
        }
      ],
      tool_choice: { type: "tool", name: toolName }
    };

    let response;
    let delay = 1500;
    const maxRetries = 2; // Low retries per model since we have fallbacks
    let isModelNotFoundError = false;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        response = await fetch("https://api.anthropic.com/v1/messages", {
          method: "POST",
          headers: {
            "x-api-key": apiKey,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json"
          },
          body: JSON.stringify(requestBody)
        });

        if (response.ok) {
          const result = await response.json();
          const toolUse = result.content?.find(c => c.type === "tool_use");
          if (!toolUse || !toolUse.input) {
            throw new Error("Claude failed to return structured tool output.");
          }
          return toolUse.input; // Success!
        }

        const errText = await response.text();
        
        // If the model is not found/not supported for this key/account, proceed to fallback
        if (response.status === 404 || errText.includes("not_found_error") || errText.includes("model_not_found")) {
          isModelNotFoundError = true;
          lastError = new Error(`Model ${model} not supported: ${errText}`);
          console.warn(`Model ${model} is not supported/found. Response: ${errText}`);
          break;
        }

        throw new Error(`Anthropic API error (${response.status}): ${errText}`);
      } catch (err) {
        if (isModelNotFoundError) {
          lastError = err;
          break;
        }
        if (attempt === maxRetries) {
          lastError = err;
          break;
        }
        console.warn(`Claude API attempt ${attempt} failed for model ${model} (${err.message}). Retrying in ${delay}ms...`);
        await new Promise(resolve => setTimeout(resolve, delay));
        delay *= 2;
      }
    }

    if (isModelNotFoundError) {
      continue; // Skip to next model
    }

    // If it was another error (like quota or billing), stop and throw immediately
    if (lastError) {
      throw lastError;
    }
  }

  throw new Error(`All Anthropic models failed. Last error: ${lastError ? lastError.message : "unknown error"}`);
}

/**
 * Extracts values for the requested headers from the parsed document markdown.
 * Sanitizes header names to meet Anthropic's parameter name requirements (no spaces).
 */
export async function extractFieldsFromText(documentText, headers, customPrompt = "") {
  if (!headers || headers.length === 0) {
    return {};
  }

  // Create clean property keys matching ^[a-zA-Z0-9_.-]{1,64}$
  const headerMap = {}; // originalHeader -> sanitizedKey
  const reverseMap = {}; // sanitizedKey -> originalHeader

  headers.forEach((header, idx) => {
    // Replace spaces and special characters with underscores, keep alphanumeric/dots/dashes
    let sanitized = header.replace(/[^a-zA-Z0-9_.-]/g, "_");
    if (sanitized.length > 50) {
      sanitized = sanitized.substring(0, 50);
    }
    // Handle duplicates
    let uniqueSanitized = sanitized;
    let counter = 1;
    while (reverseMap[uniqueSanitized]) {
      uniqueSanitized = `${sanitized}_${counter}`;
      counter++;
    }
    
    headerMap[header] = uniqueSanitized;
    reverseMap[uniqueSanitized] = header;
  });

  let prompt = `
You are a highly accurate data extraction system. Your task is to extract the exact values for the requested fields from the following document text.
If a field is not present in the document, return an empty string "" for that field.
Do not make up values. Only extract what is present or clearly implied.

Fields to extract:
`;

  headers.forEach(h => {
    prompt += `- ${h} (corresponds to the schema property name: "${headerMap[h]}")\n`;
  });

  if (customPrompt && customPrompt.trim()) {
    prompt += `
User Custom Instructions:
CRITICAL: Please strictly follow these custom user instructions when performing the extraction and mapping:
"${customPrompt.trim()}"
`;
  }

  prompt += `
Document Text:
---
${documentText}
---
`;

  const schemaProperties = {};
  for (const header of headers) {
    const sKey = headerMap[header];
    schemaProperties[sKey] = {
      type: "string",
      description: `Value of the field '${header}' extracted from the text. Empty string if not found.`
    };
  }

  const responseSchema = {
    type: "object",
    properties: schemaProperties,
    required: Object.values(headerMap)
  };

  try {
    const rawResult = await callClaude(prompt, responseSchema, "extract_fields");
    
    // Map sanitized keys back to original headers
    const result = {};
    for (const header of headers) {
      const sKey = headerMap[header];
      result[header] = rawResult[sKey] !== undefined ? String(rawResult[sKey]).trim() : "";
    }
    return result;
  } catch (error) {
    console.error("Claude field extraction error:", error);
    const fallback = {};
    for (const h of headers) {
      fallback[h] = "";
    }
    return fallback;
  }
}

/**
 * Converts unstructured PDF markdown text (e.g. parsed via LlamaParse) into structured tabular headers and rows.
 * Uses structured tool calling with Claude to guarantee format.
 */
export async function convertPdfMarkdownToTable(markdownText) {
  const prompt = `
You are an expert structured data parser. Your task is to identify and extract the main tabular data or record list from the following document text/markdown.
Convert it into a structured format containing:
1. "headers": An array of strings representing the column names of the table.
2. "rows": An array of objects where each object represents a row in the table, with keys exactly matching the column headers.

Instructions:
- If there are multiple tables, extract the primary one containing the main records (e.g. invoice list, transaction ledger, client records).
- Normalize column names to be concise and clear.
- Ensure every row object has the keys defined in "headers". If a cell is blank in a row, use empty string "".

Document Text:
---
${markdownText}
---
`;

  const responseSchema = {
    type: "object",
    properties: {
      headers: {
        type: "array",
        items: { type: "string" },
        description: "List of table column names"
      },
      rows: {
        type: "array",
        items: {
          type: "object",
          additionalProperties: { type: "string" },
          description: "Row data representing records"
        },
        description: "Array of row objects matching the headers"
      }
    },
    required: ["headers", "rows"]
  };

  try {
    return await callClaude(prompt, responseSchema, "parse_table");
  } catch (error) {
    console.error("Claude markdown to table conversion error:", error);
    throw new Error(`Failed to extract structured data from PDF via Claude: ${error.message}`);
  }
}

/**
 * Inspects parsed markdown text and dynamically discovers 4 to 6 key headers.
 */
export async function detectSchemaFromMarkdown(markdown) {
  const prompt = `
You are an expert data architect. Your task is to inspect the parsed text/markdown of a financial/transactional document (such as an invoice, receipt, bank statement, or transaction log) and identify the 4 to 6 most important, common field headers that should be extracted from this document to represent its records (e.g. Invoice ID, Date, Total Amount, Customer Name, Vendor Name, Quantity, Description, etc.).
Do not include metadata like file headers or system timestamps. Focus on headers that would be useful for comparison in a spreadsheet.

Return a JSON array of strings containing ONLY the detected field headers (between 4 and 6 headers).
For example: ["Invoice ID", "Date", "Total Amount", "Customer Name"]
`;

  const responseSchema = {
    type: "array",
    items: { type: "string" },
    description: "List of 4 to 6 key headers detected in the document."
  };

  try {
    const promptWithText = `${prompt}\n\nDocument Text:\n---\n${markdown.substring(0, 8000)}\n---`;
    return await callClaude(promptWithText, responseSchema, "detect_schema");
  } catch (error) {
    console.error("Claude schema detection error:", error);
    // Return a default schema if it fails
    return ["Invoice ID", "Date", "Amount", "Vendor"];
  }
}

/**
 * Maps PDF headers to Excel headers automatically using Claude.
 */
export async function generateAutoMapping(pdfHeaders, excelHeaders) {
  const prompt = `
You are a database integration helper. Your task is to automatically map columns from a parsed PDF dataset (pdfHeaders) to columns in a spreadsheet dataset (excelHeaders) based on name similarity and semantic meaning.

For each PDF header in the list, identify the single best matching Excel header. If there is no reasonable match, map it to an empty string "".
Do not invent Excel headers. Only select from the provided excelHeaders list or empty string.

PDF Headers:
${JSON.stringify(pdfHeaders)}

Excel Headers:
${JSON.stringify(excelHeaders)}

Return a JSON object mapping each PDF header to the matching Excel header.
For example: {"Invoice ID": "Invoice No", "Date": "Billing Date", "Total Amount": ""}
`;

  const responseSchema = {
    type: "object",
    additionalProperties: { type: "string" },
    description: "Mapping of PDF headers to Excel headers."
  };

  try {
    return await callClaude(prompt, responseSchema, "auto_map_columns");
  } catch (error) {
    console.error("Claude auto column mapping error:", error);
    // Return a simple manual fallback: match case-insensitive names
    const mapping = {};
    for (const pdfH of pdfHeaders) {
      const match = excelHeaders.find(eh => eh.toLowerCase().trim() === pdfH.toLowerCase().trim());
      mapping[pdfH] = match || "";
    }
    return mapping;
  }
}

/**
 * Generates a professional financial auditor-style Executive Reconciliation Report in Markdown.
 */
export async function generateReconciliationExecutiveSummary(comparisonResult, mappings) {
  const summary = comparisonResult.summary;
  const mismatches = comparisonResult.records
    .filter(r => r.status === "MISMATCH")
    .slice(0, 10) // pass a sample of up to 10 mismatches
    .map(r => {
      const diffs = [];
      for (const key of Object.keys(r.fields)) {
        const field = r.fields[key];
        if (field.status === "MISMATCH") {
          diffs.push(`- ${key}: PDF has "${field.pdfValue}", Excel has "${field.excelValue}"`);
        }
      }
      return `PDF Record Index ${r.pdfRowIndex} (Key Value: ${r.pdfRow ? r.pdfRow[Object.keys(mappings)[0]] || "N/A" : "N/A"}):\n${diffs.join("\n")}`;
    });

  const prompt = `
You are a senior financial auditor and data analyst. Write a professional, comprehensive Executive Reconciliation Report based on the following comparison results between a parsed PDF document and an Excel ledger sheet.

Reconciliation Summary:
- Total PDF Records parsed: ${summary.totalPdfsProcessed}
- Fully/Partially Matched Records: ${summary.totalMatches}
- Mismatching Records: ${summary.totalMismatches}
- PDF Records with No Excel Row Match: ${summary.totalNoExcelMatch}
- Excel Rows with No PDF Match: ${summary.totalUnmatchedExcelRows}

Mappings Used (PDF -> Excel):
${JSON.stringify(mappings, null, 2)}

Sample Mismatch Details:
${mismatches.length > 0 ? mismatches.join("\n\n") : "No field-level mismatches found!"}

Please write the report in Markdown. Use a professional, clean, corporate tone suitable for senior executives at MultiBank Group. Include the following sections:
1. Executive Summary: High-level overview of the audit status and overall reconciliation rate.
2. Key Findings: Detail major findings, discrepancy count, and unmatched records. Include a brief summary of why mismatches occurred (e.g. formatting differences, pricing variances).
3. Action Items: Clear, bulleted list of recommended next steps to resolve the discrepancies.
4. Audit Status: A clean, distinct status badge/section (e.g., PASSED, REVIEW REQUIRED, or CRITICAL DISCREPANCIES FOUND).

Keep lines relatively short and format beautifully. Do not include any meta-text outside of the Markdown report itself.
`;

  const responseSchema = {
    type: "object",
    properties: {
      reportMarkdown: {
        type: "string",
        description: "The complete reconciliation executive report formatted in standard Markdown."
      }
    },
    required: ["reportMarkdown"]
  };

  try {
    const result = await callClaude(prompt, responseSchema, "generate_executive_summary");
    return result.reportMarkdown;
  } catch (error) {
    console.error("Claude executive summary generation error:", error);
    return `
# Executive Reconciliation Report

**Date:** ${new Date().toLocaleDateString()}
**Status:** AUDIT COMPLETED (Error generating summary via AI)

## 1. Executive Summary
The reconciliation process has completed.
- Total PDF Records parsed: ${summary.totalPdfsProcessed}
- Fully/Partially Matched Records: ${summary.totalMatches}
- Mismatching Records: ${summary.totalMismatches}
- PDF Records with No Excel Row Match: ${summary.totalNoExcelMatch}
- Excel Rows with No PDF Match: ${summary.totalUnmatchedExcelRows}

Please review the discrepancy details in the grid below.
`;
  }
}


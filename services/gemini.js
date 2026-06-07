import { GoogleGenAI } from "@google/genai";
import dotenv from "dotenv";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
dotenv.config({ path: path.resolve(__dirname, "../.env") });

function getAiClient() {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    throw new Error("Gemini API key is not configured.");
  }
  return new GoogleGenAI({
    apiKey: apiKey,
  });
}

/**
 * Extracts values for the requested headers from the parsed document markdown.
 * Uses structured JSON output from Gemini to guarantee format.
 * 
 * @param {string} documentText - The text or markdown parsed from the PDF
 * @param {string[]} headers - The headers we want to extract
 * @param {string} customPrompt - Custom user prompt instruction to guide extraction
 * @returns {Promise<Record<string, string>>} - Map of header -> extracted value
 */
export async function extractFieldsFromText(documentText, headers, customPrompt = "") {
  const ai = getAiClient();

  if (!headers || headers.length === 0) {
    return {};
  }

  // Construct a prompt explaining the extraction task
  let prompt = `
You are a highly accurate data extraction system. Your task is to extract the exact values for the requested fields from the following document text.
If a field is not present in the document, return an empty string "" for that field.
Do not make up values. Only extract what is present or clearly implied.

Fields to extract:
${headers.map(h => `- ${h}`).join("\n")}
`;

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

  // Build schema dynamically
  const schemaProperties = {};
  for (const header of headers) {
    // Sanitize property name (must match JSON object key requirements, but we want it to map to the original header)
    // We will use the original headers since the API accepts keys with spaces.
    // If keys with spaces cause issues, we will handle them. But standard JSON schema objects allow any string as property name.
    schemaProperties[header] = {
      type: "STRING",
      description: `Value of the field '${header}' extracted from the text. Empty string if not found.`
    };
  }

  const responseSchema = {
    type: "OBJECT",
    properties: schemaProperties,
    required: headers
  };

  try {
    // Retry wrapper for robust API execution
    let response;
    let delay = 1500;
    const maxRetries = 3;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        response = await ai.models.generateContent({
          model: "gemini-2.0-flash", // stable, fast, production-ready model
          contents: [prompt],
          config: {
            responseMimeType: "application/json",
            responseSchema: responseSchema,
            temperature: 0.1,
          }
        });
        break; // success, break retry loop
      } catch (err) {
        if (attempt === maxRetries) {
          throw err; // throw on final failure
        }
        console.warn(`Gemini API attempt ${attempt} failed (${err.message}). Retrying in ${delay}ms...`);
        await new Promise(resolve => setTimeout(resolve, delay));
        delay *= 2; // exponential backoff
      }
    }

    const responseText = response.text || response.candidates?.[0]?.content?.parts?.[0]?.text;
    if (!responseText) {
      throw new Error("Empty response from Gemini API");
    }

    const parsedData = JSON.parse(responseText.trim());
    return parsedData;
  } catch (error) {
    console.error("Gemini field extraction error after retries:", error);
    // Fallback: Attempt simple text parsing or return empty object
    const fallback = {};
    for (const h of headers) {
      fallback[h] = "";
    }
    return fallback;
  }
}

/**
 * Converts unstructured PDF markdown text (e.g. parsed via LlamaParse) into structured tabular headers and rows.
 * Uses structured JSON output from Gemini to guarantee format.
 * 
 * @param {string} markdownText - Unstructured document text/markdown
 * @returns {Promise<{headers: string[], rows: Record<string, string>[]}>}
 */
export async function convertPdfMarkdownToTable(markdownText) {
  const ai = getAiClient();

  const prompt = `
You are an expert structured data parser. Your task is to identify and extract the main tabular data or record list from the following document text/markdown.
Convert it into a structured JSON object containing:
1. "headers": An array of strings representing the column names of the table.
2. "rows": An array of objects where each object represents a row in the table, with keys exactly matching the column headers.

Instructions:
- If there are multiple tables, extract the primary one containing the main records (e.g. invoice list, transaction ledger, client records).
- Normalize column names to be concise and clear.
- Ensure every row object has the keys defined in "headers". If a cell is blank in a row, use empty string "".
- Return ONLY the raw JSON object conforming to the schema. Do not put markdown code fences around the JSON.
`;

  const promptContent = `
${prompt}

Document Text:
---
${markdownText}
---
`;

  const responseSchema = {
    type: "OBJECT",
    properties: {
      headers: {
        type: "ARRAY",
        items: { type: "STRING" },
        description: "List of table column names"
      },
      rows: {
        type: "ARRAY",
        items: {
          type: "OBJECT",
          additionalProperties: { type: "STRING" },
          description: "Row data representing records"
        },
        description: "Array of row objects matching the headers"
      }
    },
    required: ["headers", "rows"]
  };

  try {
    let response;
    let delay = 1500;
    const maxRetries = 3;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        response = await ai.models.generateContent({
          model: "gemini-2.0-flash",
          contents: [promptContent],
          config: {
            responseMimeType: "application/json",
            responseSchema: responseSchema,
            temperature: 0.1,
          }
        });
        break;
      } catch (err) {
        if (attempt === maxRetries) {
          throw err;
        }
        console.warn(`Gemini PDF parse attempt ${attempt} failed (${err.message}). Retrying in ${delay}ms...`);
        await new Promise(resolve => setTimeout(resolve, delay));
        delay *= 2;
      }
    }

    const responseText = response.text || response.candidates?.[0]?.content?.parts?.[0]?.text;
    if (!responseText) {
      throw new Error("Empty response from Gemini API");
    }

    const parsedData = JSON.parse(responseText.trim());
    return parsedData;
  } catch (error) {
    console.error("Gemini markdown to table conversion error:", error);
    throw new Error(`Failed to extract structured data from PDF: ${error.message}`);
  }
}

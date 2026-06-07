import dotenv from "dotenv";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
dotenv.config({ path: path.resolve(__dirname, "../.env") });

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
 * Calls the Anthropic Claude API using native fetch with tool choice forcing.
 * This guarantees a structured JSON output conforming to the responseSchema.
 */
async function callClaude(prompt, responseSchema, toolName = "extract_data") {
  const apiKey = process.env.ANTHROPIC_API_KEY || process.env.GEMINI_API_KEY;
  if (!apiKey) {
    throw new Error("Anthropic API key is not configured.");
  }

  const normalizedSchema = normalizeSchema(responseSchema);
  
  // List of models to try in order of preference
  const models = [
    "claude-3-5-sonnet-latest",
    "claude-3-5-sonnet-20241022",
    "claude-3-5-sonnet-20240620",
    "claude-3-5-haiku-latest",
    "claude-3-5-haiku-20241022",
    "claude-3-haiku-20240307"
  ];

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
 * Uses structured tool calling with Claude to guarantee format.
 */
export async function extractFieldsFromText(documentText, headers, customPrompt = "") {
  if (!headers || headers.length === 0) {
    return {};
  }

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

  const schemaProperties = {};
  for (const header of headers) {
    schemaProperties[header] = {
      type: "string",
      description: `Value of the field '${header}' extracted from the text. Empty string if not found.`
    };
  }

  const responseSchema = {
    type: "object",
    properties: schemaProperties,
    required: headers
  };

  try {
    return await callClaude(prompt, responseSchema, "extract_fields");
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

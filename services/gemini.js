import { GoogleGenAI } from "@google/genai";
import dotenv from "dotenv";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
dotenv.config({ path: path.resolve(__dirname, "../.env") });

const apiKey = process.env.GEMINI_API_KEY;

if (!apiKey) {
  console.warn("WARNING: GEMINI_API_KEY is not set in .env file.");
}

const ai = new GoogleGenAI({
  apiKey: apiKey || "",
});

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
  if (!apiKey) {
    throw new Error("Gemini API key is not configured.");
  }

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

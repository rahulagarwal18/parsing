import { LlamaCloud } from "@llamaindex/llama-cloud";
import fs from "fs";
import dotenv from "dotenv";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
dotenv.config({ path: path.resolve(__dirname, "../.env"), override: true });

function getLlamaClient() {
  const apiKey = process.env.LLAMA_CLOUD_API_KEY;
  if (!apiKey) {
    throw new Error("LlamaCloud API key is not configured.");
  }
  return new LlamaCloud({
    apiKey: apiKey,
  });
}

/**
 * Parses a PDF file using LlamaParse and returns its full markdown text.
 * @param {string} filePath - Path to the PDF file on disk
 * @returns {Promise<{ markdown: string, text: string }>}
 */
export async function parsePdf(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }

  const client = getLlamaClient();

  // Upload file
  const fileObj = await client.files.create({
    file: fs.createReadStream(filePath),
    purpose: "parse",
  });

  // Submit and wait for parsing completion
  const result = await client.parsing.parse({
    file_id: fileObj.id,
    tier: "agentic",
    version: "latest",
    expand: ["markdown_full", "text_full"],
  });

  return {
    markdown: result.markdown_full || "",
    text: result.text_full || ""
  };
}

import express from "express";
import app from "./api/index.js";
import dotenv from "dotenv";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
dotenv.config({ path: path.resolve(__dirname, ".env"), override: true });

const PORT = process.env.PORT || 3050;

app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});

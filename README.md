# ReconcileAI - PDF & Excel Data Comparison App

ReconcileAI is a premium, web-based audit tool to automate the parsing, extraction, and comparison of PDF documents against spreadsheet records. The system utilizes LlamaParse for high-quality document extraction and the Google Gemini API for structured field parsing, presenting comparison results in an interactive frontend dashboard and generating reports in Excel, PDF, or Word.

## Features
- **Spreadsheet Parsing**: Upload Excel sheet (`.xlsx`, `.xls`) to preview rows.
- **Dynamic Field Configuration**: Choose columns to reconcile, and select the unique match key to link spreadsheets to PDFs.
- **Batch PDF Processing**: Concurrently process up to 200 PDFs with real-time SSE progress updates.
- **LlamaParse + Gemini Analysis**: Standardize layout parsing and use Google Gen AI for structured data extraction.
- **Full Reconciliation Logs**: Collapsible results table detailing match status and field discrepancies.
- **Document Exporting**: Download detailed summary reports in Excel, PDF, or Word.

---

## 🛠️ Local Setup

1. **Clone project & Install dependencies**:
   ```bash
   npm install
   ```
2. **Setup environment variables**:
   Create a `.env` file in the root directory:
   ```env
   LLAMA_CLOUD_API_KEY=your_llamaparse_key_here
   GEMINI_API_KEY=your_gemini_key_here
   PORT=3050
   ```
3. **Start the local server**:
   ```bash
   npm run dev
   ```
4. **Access the application**:
   Open [http://localhost:3050](http://localhost:3050) in your browser.

---

## 🚀 Deploying to Vercel

This application is ready to deploy on **Vercel** as a serverless project.

### Step 1: Create a GitHub Repository & Push Code
You can initialize Git and push the code:
```bash
git init
git add .
git commit -m "Initial commit of ReconcileAI"
git branch -M main
git remote add origin https://github.com/rahulagarwal18/parsing.git
git push -u origin main
```

### Step 2: Import Project to Vercel
1. Log into your Vercel Dashboard and click **Add New** > **Project**.
2. Select your repository `parsing`.
3. Under **Environment Variables**, add:
   - `LLAMA_CLOUD_API_KEY`
   - `GEMINI_API_KEY`
4. Click **Deploy**. Vercel will automatically configure the routing and serverless Node.js builders.

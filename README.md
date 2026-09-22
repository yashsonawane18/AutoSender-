# ⚡ AETHERIA AutoSend Pro — WhatsApp Automation Web Suite

> **Smart Number Normalizer, Multi-Format File Uploader (Excel, CSV, PDF, Word), Big Data Batch Controller & Persistent IndexedDB Storage — 100% Client-Side.**

🌐 **Live Demo Hosted on GitHub Pages:**  
👉 **[https://yashsonawane18.github.io/AutoSender-/](https://yashsonawane18.github.io/AutoSender-/)**

---

## 🌟 Key Features

- **Multi-Format Document Upload**:
  - Excel (`.xlsx`, `.xls`) multi-sheet aware via SheetJS.
  - CSV (`.csv`) with automatic delimiter detection.
  - PDF (`.pdf`) table and text extraction via PDF.js.
  - Word (`.docx`) table & paragraph parsing via Mammoth.js.
- **Smart Phone Number Normalizer**:
  - Automatically fixes Excel scientific notation (e.g. `9.87654E+09` -> `9876543210`).
  - Cleans float decimals (`9876543210.0` -> `9876543210`).
  - Handles leading trunk `0`s, international `+` codes, hyphens, and spaces.
  - Resolves 10-digit numbers starting with `91` without conflict.
- **Big Data Starting Sr. No. & Batch Control**:
  - Custom starting position (`Start from Sr. No.` / `End at Sr. No.`).
  - 1-Click `⏩ Jump to 1st Unsent` contact.
  - Quick batch presets: `[+25]`, `[+50]`, `[+100]`, `[All]`.
  - Individual table row `[🎯 Start Here]` action.
- **Zero-Ban WhatsApp Dispatch Engine**:
  - **Mode A (Fast 1-Click Queue)**: Zero popup blocker issues, keyboard shortcut (`Spacebar`/`Enter`) supported.
  - **Mode B (Auto-Loop Runner)**: Single-tab re-use (`AetheriaWhatsAppTab`) with customizable pacing delays (3s–30s).
- **Persistent Database (IndexedDB)**:
  - Automatically saves campaigns and contact sheets in the browser.
  - Resume any campaign from history with 1 click.
  - Pre-built and customizable Template Library.
  - Full database JSON backup and restore.

---

## 🚀 How to Run Locally

Simply clone the repository and open `index.html` in any browser:
```bash
git clone https://github.com/yashsonawane18/AutoSender-.git
cd AutoSender-
# Open index.html in your browser or run:
npx serve .
```

---

## 🔒 Security & Privacy

This application is **100% client-side**. No contact data, spreadsheets, or message contents are ever sent to any third-party server or backend. All data stays strictly in your browser.

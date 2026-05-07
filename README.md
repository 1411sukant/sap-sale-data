# 📊 Sales Register PDF → Monthly Excel

A **Streamlit** web app that parses SAP Business One Sales Register PDFs and produces a formatted, month-wise Excel report.

---

## 🚀 Live Demo (Streamlit Cloud)

[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://your-app-name.streamlit.app)

> Replace the badge URL after you deploy (Step 4 below).

---

## ✨ Features

| Feature | Detail |
|---|---|
| 📤 PDF Upload | Drag-and-drop or browse · max 50 MB |
| 📅 Month-wise grouping | Based on **Invoice Date** column · FY order Apr → Mar |
| ✅ Sale / Service Inv | Treated as **positive** — added to totals |
| ↩️ Credit Note | Treated as **negative** — subtracted from totals |
| 📊 Extracted columns | Sales Value · SGST · CGST · IGST |
| 📥 Excel output | 3 formatted sheets: Monthly Summary · Transactions · Quarter View |

---

## 🗂️ Project Structure

```
sales-register-extractor/
├── app.py                  ← Main Streamlit application
├── requirements.txt        ← Python dependencies
├── .streamlit/
│   └── config.toml         ← Theme & server settings
├── .gitignore
└── README.md
```

---

## 🛠️ Local Setup

```bash
# 1. Clone the repo
git clone https://github.com/YOUR_USERNAME/sales-register-extractor.git
cd sales-register-extractor

# 2. Create a virtual environment (recommended)
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run locally
streamlit run app.py
# Opens at http://localhost:8501
```

---

## ☁️ Deploy to Streamlit Community Cloud (Free)

### Step 1 — Push to GitHub

```bash
git init
git add .
git commit -m "Initial commit: Sales Register Extractor"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/sales-register-extractor.git
git push -u origin main
```

### Step 2 — Sign in to Streamlit Cloud

Go to **[share.streamlit.io](https://share.streamlit.io)** and sign in with your GitHub account.

### Step 3 — Create a New App

1. Click **"New app"**
2. Select your repository: `YOUR_USERNAME/sales-register-extractor`
3. Branch: `main`
4. Main file path: `app.py`
5. Click **"Deploy!"**

### Step 4 — Update README badge

Once deployed, copy your app URL (e.g. `https://sales-register.streamlit.app`) and update the badge at the top of this README.

---

## 📐 Processing Logic

```
For each transaction in the PDF:
    ├── Invoice Date → determines the month bucket
    ├── Doc Type == "Sale" or "Service" → positive values
    └── Doc Type == "CREDIT" → negative values (subtracted)

Monthly totals = Σ(Sale invoices) − Σ(Credit notes)
```

---

## 📋 Excel Output (3 Sheets)

### Sheet 1 — Monthly Summary
Month-wise totals for Sales Value, SGST, CGST, IGST, Total Tax, and Gross Total — with counts of invoices and credit notes per month.

### Sheet 2 — Transactions
Full detail log of every parsed transaction: doc number, type, invoice date, month, and all tax amounts. Credit notes highlighted in red.

### Sheet 3 — Quarter View
Q1 (Apr–Jun) · Q2 (Jul–Sep) · Q3 (Oct–Dec) · Q4 (Jan–Mar) aggregated totals.

---

## 🐛 Issues / Contributions

Found a bug or want to improve parsing for other SAP formats? Open an issue or submit a PR!

---

## 📄 License

MIT — free to use and modify.

# 📊 Macro Data Extractor

A Python tool that extracts selected macroeconomic indicators (e.g. Call Money Rate, Monthly Exports, Remittances, Reserve Money, Excess Liquidity) from a large historical Excel dataset for a chosen date range — with a clean web interface to select data, visualize it with bar and line charts, and export it to a new Excel file in the original layout.

Built as a way to replace a fully manual, error-prone Excel workflow with a repeatable, self-serve tool.

---

## ✨ Features

- **Interactive web interface** built with [Streamlit](https://streamlit.io/) — no coding needed to use it
- **Multi-select indicator picker** — choose one or many macro indicators at once
- **Flexible date range filtering** — pick any "From" / "To" month across the full dataset
- **In-app data visualization** — bar charts and line charts rendered instantly, no extra file needed
- **One-click Excel export** — outputs a new `.xlsx` file matching the original sheet's layout
- **Terminal version included** — a lightweight CLI script for automation/scripting use cases
- **Self-updating** — new indicators or months added to the source file are automatically picked up, no code changes required

---

## 🖥️ Screenshots

<!-- 📸 Add a screenshot of the full interface here -->

<!-- 📸 Add a screenshot of a generated bar/line chart here -->

---

## 🛠️ Tech Stack

- **Python 3**
- **[Streamlit](https://streamlit.io/)** — web interface
- **[openpyxl](https://openpyxl.readthedocs.io/)** — Excel file reading/writing
- **[Matplotlib](https://matplotlib.org/)** — charting

---

## 🚀 Getting Started

### Prerequisites
- Python 3.10+
- pip

### Installation

```bash
git clone https://github.com/salamander55/Macro_Extractor-Final-.git
cd Macro_Extractor-Final-
pip install openpyxl streamlit matplotlib
```

> **Note:** This repo does not include the source data file (`monthly_macro.xlsx`) since it contains proprietary data. To run this project, supply your own Excel file with a `"Monthly"` sheet in the same structure (indicator names in column A, units in column B, monthly dates from column D onward).

### Running the app

```bash
streamlit run app.py
```

This opens the interface in your browser at `http://localhost:8501`.

### Running the terminal-only version

```bash
python macro_data_extractor.py
```

---

## 📂 Project Structure

```
├── app.py                     # Streamlit web interface
├── macro_data_extractor.py    # Core extraction logic + terminal version
├── run_app.py                 # Launcher for running via an IDE's Run button
└── README.md
```

---

## 📌 Notes

- Charts are generated live in the browser and are not saved into the exported Excel file.
- Every export is timestamped, so previous outputs are never overwritten.
- Months or indicator rows added to the source file later are picked up automatically — no code changes needed.

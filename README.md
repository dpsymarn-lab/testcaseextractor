# CBAM Test Case Extractor

Extracts E2E test cases from CBAM `.docx` specification documents and produces a **Jira Zephyr Scale**-compatible CSV file ready to import directly into your project.

---

## What it does

- Reads one or more `.docx` (or `.txt` / `.md`) files containing CBAM test case tables
- Finds every test case heading matching the pattern `CBAM-…-TC-NNNN`
- Extracts: ID, Name, Purpose, Actors, Pre-Conditions, Post-Conditions, Pass Criteria, and all test steps (action / data / expected result)
- Writes a CSV in the format expected by **Jira Zephyr Scale** bulk import (one row per step, metadata on the first step row only)

---

## Files

| File | Description |
|---|---|
| `extract_test_cases.py` | Core extraction logic. Can also be run from the command line. |
| `app_ui.py` | Desktop GUI — the recommended way to use the tool. |
| `requirements.txt` | Python dependencies. |

---

## Requirements

### Python

Python **3.8 or later** is required.

Download from https://www.python.org/downloads/ and make sure to tick **"Add Python to PATH"** during installation on Windows.

Verify your installation:

```bash
python --version
# or on macOS/Linux:
python3 --version
```

### Python packages

Install dependencies with:

```bash
pip install -r requirements.txt
```

This installs:

| Package | Purpose |
|---|---|
| `python-docx` | Reads `.docx` Word documents |

> `tkinter` (used by the GUI) ships with the standard Python installation on Windows and macOS. On Linux you may need to install it separately — see the note below.

---

## Installation

1. **Download** or clone this folder to your computer.
2. Open a terminal (Command Prompt / PowerShell on Windows, Terminal on macOS/Linux).
3. Navigate to the folder:

```bash
cd path/to/cbam-extractor
```

4. Install dependencies:

```bash
pip install -r requirements.txt
```

---

## Running the GUI (recommended)

```bash
python app_ui.py
# or on macOS/Linux:
python3 app_ui.py
```

### Steps inside the app

1. Click **+ Add Files** and select one or more `.docx` files.
2. The **Output file** field defaults to `test_cases_zephyr.csv` in your current directory. Click **Browse…** to change the location.
3. Click **⚡ Extract Test Cases**.
4. The log panel shows progress. When done, a summary dialog appears and the CSV is saved.

---

## Running from the command line (advanced)

```bash
# Single file, default output name
python extract_test_cases.py document.docx

# Multiple files
python extract_test_cases.py doc1.docx doc2.docx

# Custom output path
python extract_test_cases.py doc1.docx doc2.docx -o my_output.csv

# Works with pre-converted text files too
python extract_test_cases.py doc1.txt doc2.md
```

---

## CSV output format

The output CSV is formatted for **Jira Zephyr Scale** bulk import. Each test case occupies multiple rows — one per step — with metadata only on the first row.

| Column | Description |
|---|---|
| `Name` | Test case title (first step row only) |
| `ID` | Full CBAM test case ID (first step row only) |
| `Source Document` | Name of the source `.docx` file |
| `Objective` | Purpose / objective of the test case |
| `Precondition` | Pre-conditions |
| `Labels` | Set to `E2E` for all test cases |
| `Priority` | Default: `Medium` |
| `Status` | Default: `Draft` |
| `Step` | Step number |
| `Step Action` | What the tester does |
| `Step Data` | Test data for that step |
| `Step Expected Result` | Expected system response |
| `Postcondition` | Post-conditions |
| `Pass Criteria` | Pass criteria |

### Importing into Jira Zephyr Scale

1. In Jira, go to **Zephyr Scale → Test Cases**.
2. Click **Import** and select **CSV**.
3. Upload the generated CSV file.
4. Map columns as prompted (column names match Zephyr Scale defaults, so mapping is usually automatic).

---

## Linux note — tkinter

On some Linux distributions, `tkinter` is not bundled with Python and must be installed separately:

```bash
# Debian / Ubuntu
sudo apt-get install python3-tk

# Fedora
sudo dnf install python3-tkinter

# Arch
sudo pacman -S tk
```

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| `ModuleNotFoundError: No module named 'docx'` | Run `pip install python-docx` |
| `ModuleNotFoundError: No module named 'tkinter'` | See the Linux note above |
| `No test cases found` | Check that your document contains headings matching `CBAM-…-TC-NNNN` |
| CSV opens garbled in Excel | Open Excel → Data → From Text/CSV and choose **UTF-8** encoding |

# Run workspace

This folder is the **working directory**. Launch the pipeline from here.

---

## Usage
1) Place **exactly one** `.xlsx` here (your input workbook).  
2) Create/activate a virtual environment and install deps:
   ```bash
   python -m venv .venv
   # Windows:
   .venv\Scripts\activate
   # macOS/Linux:
   source .venv/bin/activate
   pip install -r ../requirements.txt
   ```
3) Run:
   ```bash
   python summary.py
   ```

Result: the Excel workbook in this folder gets its **Summary** sheet updated, and a new Word file
`3.2_analysis_of_liquidity_ratios.docx` is created here.

---

## Important
- **Close the Excel workbook** before running. If it is open, Windows will lock the file and the script cannot write the Summary sheet.
- **Make sure there is _no_ Word file named `3.2_analysis_of_liquidity_ratios.docx` in this folder before running.**  
  The script **does not overwrite** an existing file with that name — it will raise an error.  
  If such a file exists, **delete or rename** it first (and ensure it is closed).
- Keep **only one** `.xlsx` in this folder when running. If there are multiple `.xlsx` files, the script will stop with a clear error.
- The **Summary** sheet in the workbook may be **absent or outdated** — the script **creates/rewrites** it from scratch each run (so it’s fine to keep either an empty sheet or none at all).

---

## Expected files in `run/`
```
run/
├─ summary.py                      # entry point (run from here)
├─ <your_workbook>.xlsx            # exactly one .xlsx (input)
└─ 3.2_analysis_of_liquidity_ratios.docx   # generated output (created on run; must NOT pre-exist)
```

---

## Outputs
- Updates the Excel **Summary** sheet (in-place).
- Creates **Word** section: `3.2_analysis_of_liquidity_ratios.docx` (fails if a file with this exact name already exists).

---

## Troubleshooting
- **PermissionError / “cannot open/write file”** → Close Excel/Word; ensure files are not read-only; run again.
- **“No input workbook found”** → Place exactly one `.xlsx` in this folder.
- **“Multiple .xlsx files found”** → Leave only the intended workbook.
- **“Output Word file already exists”** → Delete or rename `3.2_analysis_of_liquidity_ratios.docx` and run again.
- **“Missing sheet/columns”** → Verify required sheets in the workbook:  
  `years`, `parameters`, `income_statement`, `balance_sheet`, `ratio_norms`.

---

## Notes
- This folder is a **workspace**, not a data store — do **not** commit real client workbooks. For portfolio, use screenshots in `/assets`.
- You can re-run the script any time: it **rebuilds** the Summary on each run.
- The Word file is **created only if** `3.2_analysis_of_liquidity_ratios.docx` does **not** already exist. If it exists, delete/rename it before running.

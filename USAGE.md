# Smart Data Import — Usage Guide 📘

Three clicks: **Download Template → attach the filled file → Start Import.**

---

## Step 1: Get a template (optional but recommended)

1. Open **Smart Data Import** (sidebar → *Smart Data Import* workspace → *New Import*).
2. Save the document with a name, e.g. `Q3 Master Data Migration`.
3. Click **Download Template** in the toolbar and pick the DocType you want to import into.
   - Untick *Include optional fields* to get **only the mandatory columns** — the fastest way to a valid file.
   - Sheet 1 holds the headers to fill (mandatory columns are highlighted red). Sheet 2, *Field Guide*, lists each column's fieldname, type, whether it is mandatory, and which DocType a Link column must point to.
4. Fill in your data starting on row 2.

Already have your own file? That works too:

- Headers must be in **row 1**. Column order does not matter.
- Headers may be labels (`Customer Name`) or fieldnames (`customer_name`); spaces, underscores, hyphens and casing are all treated the same.
- Name the file or the sheet after the DocType (`Customer Group.xlsx`, `Customer.csv`) to help auto-detection.
- `.xlsx`, `.xlsm` and `.csv` are supported. `.csv` may use comma, semicolon, tab or pipe separators. Old `.xls` files must be re-saved as `.xlsx`.

---

## Step 2: Attach your files

1. In the **Data Files** table, add one row per file and attach it.
2. Leave **Target DocType** empty to have it detected, or pick it yourself — your choice is never overwritten.
3. **Sheet Name** applies to Excel files only; empty means the first sheet.
4. Save.

On save the files are analyzed automatically: rows are counted, the target DocType is detected, columns are matched, and the execution order is calculated. Each file row then shows:

- **Mapping & Warnings** — how many columns matched, which columns will be ignored, and any mandatory field that is missing from the file.
- **Result / Error Log** — why a file cannot be used (wrong type, file missing on disk, no data rows, DocType not detected).

The headline shows the execution flow, for example: `✨ 12,480 rows in this order: Customer Group (12) ➔ Customer (8,400) ➔ Sales Order (4,068) — ready to import.`

---

## Step 3: Check the mapping before importing

Click **Preview & Column Mapping**. For every file you see:

- the target DocType and row count,
- each column and the field it will be imported into,
- **columns that will be ignored** (a typo in a header shows up here),
- **mandatory fields missing from the file**,
- the first rows exactly as the engine reads them.

If columns are being ignored, click **Add Mapping Skeleton**. It fills *Column Mapping & Defaults (JSON)* with an entry per unmapped column:

```json
{
  "Customer": {
    "column_map": { "Cust Name": "" }
  }
}
```

Type the fieldname on the right-hand side, save, and preview again. The same block also sets values that are not in the file at all:

```json
{
  "Customer": {
    "column_map": { "Cust Name": "customer_name" },
    "defaults":   { "customer_type": "Company" }
  }
}
```

---

## Step 4: Choose your options (optional)

Under **⚙️ Advanced Settings**:

| Option | Use it when |
| :--- | :--- |
| **Import Mode** | `Insert New Records` for fresh data, `Update Existing Records` to edit records that already exist, `Insert and Update` to do both. Matching is by ID, then unique fields, then the title field. |
| **Skip Rows That Already Exist** | Re-running a partially completed import — existing rows are skipped and logged instead of failing. |
| **Ignore Missing Link Errors** | Linked masters will be created later. |
| **Ignore Missing Mandatory Fields** | The file legitimately lacks a required field. |
| **Stop Immediately On First Error** | A dry run where you want to see the first problem and stop. |
| **Batch Chunk Size** | Lower it (e.g. 500) for DocTypes with heavy server-side logic; raise it for simple masters. |

---

## Step 5: Run it

1. Click **🚀 Start Import**, confirm the row count and mode.
2. The job runs in a background worker — you can close the page, progress is stored on the document.
3. The progress bar, *Imported*, *Failed* and *Skipped* counters update live. The form is never marked as unsaved by these updates.

---

## Step 6: Review the results

- **Completed** — `🎉 Done: N records imported in Xs (M skipped).`
- **Partial Success** / **Failed** — the errors are shown right on the form, three ways:
  1. **A dialog opens automatically** when the import finishes with failures, grouping them by cause:

     | Rows | Type | Problem |
     | ---: | :--- | :--- |
     | 246 | Failed | `Could not find Supplier Group: …` |
     | 112 | Failed | `Supplier Type cannot be "…". It should be one of "Company", "Individual", "Partnership"` |
     | 24 | Skipped | `Already exists as …` |

     Reopen it any time with **❗ View Errors** in the toolbar, or the *See what went wrong* link in the coloured banner.
  2. **Failed & Skipped Rows** table on the form — row number, type, DocType and the exact reason for each problem row (up to 500).
  3. **Download Import Log (.xlsx)** — the complete list, including runs with more than 500 problem rows.
- If the job itself crashed, the traceback is in *Detailed Error Log*.

To run the same document again, click **Reset for New Run** (counters and logs are cleared; already-imported records are not deleted), fix your file or options, then start again. Tick *Skip Rows That Already Exist* so the rows that succeeded the first time are skipped.

---

## Step 7: Made a mistake? Roll it back ⏪

Every record an import creates is listed in a **Rollback Manifest** attached to the document, so a bad import can be undone.

1. Click **⏪ Rollback Import**.
2. The dialog shows exactly what will be deleted, grouped per DocType, e.g. `Customer 8,400 · Customer Group 12`.
3. Confirm. Deletion runs in the background with live progress; the status becomes **Rolled Back** and *Rolled Back Records* shows the count.

What it does and does not do:

- **Deletes only the records this import created.** Records that already existed are never touched.
- **Deletes newest first**, so tasks go before their projects and tree children before their parents. Self-referencing links (`parent_task`, `parent_item_group`, …) are detached first, otherwise a parent/child pair blocks its own deletion.
- **Submitted documents are cancelled first**, then deleted.
- **Records used elsewhere are kept.** If another document links to an imported record, it is skipped, listed in the *Detailed Error Log*, and left in the manifest — delete the blocker and click **Rollback Import** again to finish the job. Only tick *"Ignore link validations (dangerous)"* if you accept the broken links it can leave behind.
- **Updates cannot be undone.** Rows imported in *Update Existing Records* / *Insert and Update* mode changed existing records; their old values are not stored here. Use each document's version history for those. The dialog tells you how many such records there are.
- **Reset for New Run keeps the manifest**, so you can still undo the previous run after resetting. Once everything in the manifest is gone, the manifest is removed and the Rollback button disappears.
- Rollback requires *delete* permission on the document and on every DocType involved — checked before anything is deleted.

---

## Python API

```python
import frappe
from smart_data_import.smart_import_engine import (
    SmartImportEngine, start_smart_import, get_import_preview,
    download_import_template, reset_import,
    get_rollback_summary, rollback_import,
)

# Analyze (also happens automatically on save and on Start Import)
SmartImportEngine("SDI-2026-08-00001").analyze_files_and_build_graph()

# Inspect the mapping without writing anything
get_import_preview("SDI-2026-08-00001", sample_size=5)

# Queue the background import
start_smart_import("SDI-2026-08-00001")

# Run it synchronously (scripts, tests, bench console)
SmartImportEngine("SDI-2026-08-00001").execute_import()

# Template for a DocType, and a clean slate for a re-run
download_import_template("Customer", include_optional=0)
reset_import("SDI-2026-08-00001")

# Undo: inspect first, then queue the rollback (or run it synchronously)
get_rollback_summary("SDI-2026-08-00001")     # {'deletable': 8412, 'doctypes': [...], ...}
rollback_import("SDI-2026-08-00001")          # background
SmartImportEngine("SDI-2026-08-00001").execute_rollback()          # synchronous
SmartImportEngine("SDI-2026-08-00001").execute_rollback(force=True)  # ignore link checks
```

---

## Troubleshooting

| Symptom | Cause / fix |
| :--- | :--- |
| *"Could not detect the target DocType"* | Pick it manually in the **Target DocType** column. |
| A column is silently missing from imported records | It is listed as *ignored* in **Preview & Column Mapping** — map it via `column_map`. |
| I can't see why rows failed | Click **❗ View Errors** (or the link in the banner) for the grouped causes; every row is also in the **Failed & Skipped Rows** table. |
| *"File could not be found on disk"* | The attachment was deleted; remove the row and upload the file again. |
| *"No data rows found"* | Row 1 must be headers and data must start on row 2. |
| Rows fail with a mandatory-field error | Add the column, set it under `defaults`, or tick *Ignore Missing Mandatory Fields*. |
| `Hierarchy` entries in the log | The parent row could not be found, or the parent record breaks a business rule (e.g. ERPNext requires a parent Task to be a Group Task). |
| Status stuck on `Processing` | The background worker is not running (`bench start`). Once it fails or is restarted the document is set to `Failed` with the traceback. |
| Rollback button is missing | There is no manifest: the import ran before rollback existed, only updated records, or everything it created has already been deleted. |
| Rollback left some records behind | They are still linked from documents outside the import. The reason for each is in *Detailed Error Log* — remove the blocker and roll back again. |

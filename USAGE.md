# Smart Data Import - Comprehensive Usage Guide 📘

This guide provides end-to-end instructions for setting up, configuring, and executing high-speed batch imports using **Smart Data Import**.

---

## 🛠️ Step 1: Preparing Your Data Files

Smart Data Import supports both Excel (`.xlsx`, `.xlsm`) and `.csv` file formats.

### File Naming Best Practices
- Name your files or sheet names after the target DocType (e.g. `Customer Group.xlsx`, `Customer.csv`, `Sales Order.xlsx`).
- The engine uses filename & sheet name matching alongside column header analysis for automatic DocType detection.

### Spreadsheet Headers
- Place column headers in **Row 1** of your file or sheet.
- Column headers can be human-readable field labels (e.g., `Customer Group`, `Customer Name`) or exact field names (`customer_group`, `customer_name`).

---

## 🚀 Step 2: Creating an Import Record

1. Navigate to **Smart Data Import** list view in ERPNext / Frappe Desk.
2. Click **+ Add Smart Data Import**.
3. Enter an **Import Name / Description** (e.g. `Q3 Master Data Migration`).
4. In the **Data Files** table, attach your Excel (`.xlsx`) or CSV files.
5. Click **Save**.

> 💡 **Auto-Analysis**: Upon saving, Smart Data Import automatically inspects your files, auto-detects target DocTypes, counts rows, and builds the Directed Acyclic Graph (DAG) for tier execution order.

---

## ⚙️ Step 3: Advanced Options & Settings

Under **Advanced Settings**, configure optional import behavior:

- **Ignore Duplicates (Primary Key)** (`ignore_duplicates`):
  - When enabled, records that already exist in the database (by primary key `name` or title field) are skipped automatically.
  - Prevents `DuplicateEntryError` from stopping or failing your import batch.
- **Ignore Missing Link Errors** (`ignore_link_errors`):
  - Allows document creation even if linked records are missing.
- **Stop Immediately On First Error** (`stop_on_error`):
  - Halts processing as soon as a failure is encountered.
- **Batch Chunk Size** (`batch_size`):
  - Number of records processed per database commit transaction chunk (Default: `5,000`).

---

## ⚡ Step 4: Executing the Import

1. Click the prominent primary button **"🚀 Start Import Now"** in the document header.
2. Confirm the import execution modal prompt.
3. The import will launch asynchronously in the background.
4. **Real-time Monitoring**: Progress bar and status indicators update automatically live on screen.
5. **Auto Button Management**: Action buttons (`🚀 Start Import Now` and `🔄 Re-Analyze Files`) are automatically hidden while processing and when the import reaches `Completed` state.

---

## 📑 Step 5: Reviewing Results & Downloading Failure Logs

- When completed successfully, a green banner alert indicates overall inserted record counts.
- If any rows fail, the status will show **Partial Success** or **Failed**.
- Click **"Download Failed Rows Log (.xlsx)"** under progress section to retrieve an Excel sheet containing row numbers and exact error tracebacks.

---

## 💻 API & Python SDK Examples

### Triggering Analysis via Python
```python
import frappe

# Analyze attached files
doc = frappe.get_doc("Smart Data Import", "SDI-2026-00001")
doc.analyze_dependencies()
```

### Triggering Background Import via Python
```python
from smart_data_import.smart_import_engine import start_smart_import

# Queue background import job
start_smart_import("SDI-2026-00001")
```

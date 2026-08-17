# Smart Data Import 🚀

**Smart Data Import** is a batch data import engine for Frappe and ERPNext. It processes large datasets across multiple Excel (`.xlsx`, `.xlsm`) or CSV files with automated dependency resolution, memory-efficient streaming and detailed per-row error reporting.

> 📘 **User Guide**: See [USAGE.md](USAGE.md) for step-by-step instructions.

---

## 🌟 Key Features

### 1. 🧭 Guided, 3-Click Workflow
- **Download Template** builds a ready-to-fill `.xlsx` for any DocType: mandatory columns first (highlighted red), plus a *Field Guide* sheet documenting every column's type, whether it is required, and which DocType a Link column points to.
- **Preview & Column Mapping** shows — before a single record is written — which columns map to which fields, which columns will be *ignored*, which mandatory fields are missing, and the first rows exactly as the engine reads them.
- **Start Import** analyzes automatically if needed, so one click is enough.
- **Reset for New Run** clears counters and logs so a finished or failed import can be repeated.

### 2. ⏪ Rollback (Undo an Import)
- Every record the engine writes is appended to a **rollback manifest** (`Rollback_<DocName>.csv`, a private attachment listing `action, doctype, name`). It is written incrementally, so it costs nothing for millions of rows and survives a crash mid-import.
- **Rollback Import** deletes exactly those records — newest first, so child rows, later tiers and tree children go before the parents they point at.
- Self-referencing links (`parent_task`, `parent_item_group`, ...) are detached first, because a hierarchy links both ways and would otherwise be undeletable in *any* order. Trees are saved through the document, keeping `lft`/`rgt` intact.
- Submitted documents are cancelled before deletion; records already gone count as undone.
- Records still referenced from outside the import are **kept**, reported per record, and left in the manifest so a retry picks up just those. Link validations are only bypassed with the explicit *"Ignore link validations"* option.
- Updated records (Update / Insert-and-Update modes) are recorded but never deleted — their previous values are not recoverable, and the dialog says so.

### 3. 🔄 Automated Topological Dependency Resolution (DAG)
- Builds a Directed Acyclic Graph of the target DocTypes and uses Kahn's topological sort to compute execution tiers (e.g. `Customer Group` → `Customer` → `Sales Order`), so parents exist before the rows that reference them.
- Self-referential hierarchies (`parent_account`, `parent_item_group`, `parent_task`) are handled with a two-pass insertion. For **tree DocTypes the parent link is applied through the document**, keeping the nested set (`lft`/`rgt`) consistent; flat DocTypes take the fast direct-write path.

### 4. 🧠 DocType Auto-Detection
- A weighted score over column headers, labels, primary-key naming, sheet name and file name infers the target DocType. Custom fields are included in the index, and the whole index is built with two queries instead of one `get_meta()` per DocType.
- A DocType you select manually is **never** overwritten by auto-detection.

### 5. 🔁 Import Modes That Actually Do Something
- `Insert New Records` – create only.
- `Update Existing Records` – update matched records only; unmatched rows are logged as skipped.
- `Insert and Update` – upsert.
- Existing records are matched by ID (`name`), then by unique fields, then by the title field.

### 6. 🔑 Duplicate Handling
- **Skip Rows That Already Exist** skips duplicates instead of failing the batch, and every skipped row is recorded with its reason.

### 7. ⚡ Memory-Efficient Streaming
- Rows are streamed one at a time, inserted in configurable chunks (default `5,000`) with explicit `gc.collect()` between chunks.
- CSV delimiters (`,`, `;`, tab, `|`) are detected automatically; UTF-8 BOM files are handled.

### 8. 🎯 Real-Time Progress
- Live socket updates (`smart_import_progress`) drive the progress bar and counters. Progress writes bypass the document timestamp, so an open form is never marked dirty mid-import, and a realtime hiccup can never abort a running import.

### 9. ❗ Errors Visible in the UI
- Problem rows land in a **Failed & Skipped Rows** table on the document itself — row number, type (`Failed` / `Skipped` / `Hierarchy` / `Rollback`), DocType and the exact reason, with Frappe's HTML stripped out. The section only appears when there is something to show.
- **View Errors** (toolbar, and a link in the status headline) opens a **grouped** summary: row-specific details are collapsed into causes, so 496 failures read as `246 × Could not find Supplier Group: …`, `112 × Supplier Type cannot be "…"` instead of 496 lines.
- The dialog opens by itself the moment an import finishes with failures, and offers the full log for download.
- Up to 500 problem rows are kept on the document (`ERROR_ROWS_IN_UI`); the complete list is always in the private Excel log `Import_Log_<DocName>.xlsx`.
- If the background job crashes, the document is set to `Failed` with the traceback instead of being stuck on `Processing` forever.
- Newer fields degrade gracefully on a site where the code arrived before `bench migrate` did — an import will not crash because a column is missing.

---

## ⚙️ Configuration

| Setting | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `batch_size` | Int | `5000` | Rows per database commit chunk (minimum 500). |
| `import_type` | Select | `Insert New Records` | `Insert New Records`, `Update Existing Records`, `Insert and Update`. |
| `ignore_duplicates` | Check | `0` | Skip rows whose record already exists. |
| `ignore_link_errors` | Check | `0` | Create records even when a linked record is missing. |
| `ignore_mandatory_errors` | Check | `0` | Create records even when mandatory fields are absent. |
| `stop_on_error` | Check | `0` | Abort the whole import on the first failed row. |
| `auto_detect_doctype` | Check | `1` | Infer the target DocType for rows where it is blank. |
| `skip_empty_rows` | Check | `1` | Ignore blank spreadsheet rows. |
| `clean_whitespace` | Check | `1` | Trim leading/trailing whitespace from text values. |
| `filter_rules_json` | Code | – | Per-DocType `column_map` and `defaults` (see below). |

Read-only progress fields: `imported_records`, `failed_records`, `skipped_records`, `rolled_back_records`, `progress_percent`, `execution_time_seconds`, plus the `error_file` (import log) and `rollback_file` (rollback manifest) attachments.

### Column Mapping & Defaults

Use this when a header cannot be matched automatically, or when a value is not in the file at all:

```json
{
  "Customer": {
    "column_map": { "Cust Name": "customer_name", "Grp": "customer_group" },
    "defaults":   { "customer_type": "Company", "territory": "All Territories" }
  }
}
```

The **Preview & Column Mapping** dialog can generate this skeleton for you, pre-filled with every unmapped column. Invalid JSON or unknown options are rejected on save.

---

## 🛠️ Whitelisted API

```python
# Analyze attached files and build the execution DAG
frappe.call("smart_data_import.smart_import_engine.analyze_smart_import", {"doc_name": "SDI-2026-08-00001"})

# Queue the background import (analyzes first if needed)
frappe.call("smart_data_import.smart_import_engine.start_smart_import", {"doc_name": "SDI-2026-08-00001"})

# Column mapping + sample rows, no writes
frappe.call("smart_data_import.smart_import_engine.get_import_preview", {"doc_name": "SDI-2026-08-00001", "sample_size": 5})

# Generate an .xlsx template for a DocType
frappe.call("smart_data_import.smart_import_engine.download_import_template", {"target_doctype": "Customer", "include_optional": 0})

# Clear counters/logs for a re-run
frappe.call("smart_data_import.smart_import_engine.reset_import", {"doc_name": "SDI-2026-08-00001"})

# Failures grouped by cause, as shown in the View Errors dialog
frappe.call("smart_data_import.smart_import_engine.get_error_summary", {"doc_name": "SDI-2026-08-00001", "limit": 25})

# What an undo would delete, grouped per DocType (read-only)
frappe.call("smart_data_import.smart_import_engine.get_rollback_summary", {"doc_name": "SDI-2026-08-00001"})

# Queue the rollback; force=1 additionally bypasses link validations
frappe.call("smart_data_import.smart_import_engine.rollback_import", {"doc_name": "SDI-2026-08-00001", "force": 0})
```

Every endpoint checks document permissions (`doc.check_permission`), template generation requires `create` permission on the target DocType, and rollback requires `delete` permission on the document **and** on every DocType it would delete from (checked before anything is removed).

---

## 🧪 Testing

```bash
bench --site <site-name> set-config allow_tests true   # once, on a test site
bench --site <site-name> run-tests --app smart_data_import
```

The suite covers header normalization, column mapping (labels, fieldnames, overrides, blank columns), mandatory-field detection, DocType auto-detection, dependency tiers with self-reference, file/JSON validation, missing-file reporting, template generation, an end-to-end import including a duplicate-skipping re-run, and rollback (full undo, refusal when only updates were made, and tolerance of records deleted by hand).

---

## 📜 License

MIT License — ERPNext AI Team & Contributors.

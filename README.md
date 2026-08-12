# Smart Data Import 🚀

**Smart Data Import** is an enterprise-grade, high-performance batch data import engine for Frappe and ERPNext. It is engineered to process large datasets (millions of records) across multiple Excel (`.xlsx`) or CSV files with automated dependency resolution, memory-efficient streaming, and advanced error handling.

---

## 🌟 Key Features

### 1. 🔄 Automated Topological Dependency Resolution (DAG)
- Constructs a **Directed Acyclic Graph (DAG)** of target DocTypes using Kahn's Topological Sorting algorithm.
- Automatically calculates execution tiers (e.g. `Customer Group` -> `Customer` -> `Sales Order`) to ensure parent records are inserted before child references.
- Handles self-referential hierarchies (e.g., `parent_account`, `parent_item_group`, `parent_task`) via a two-pass insertion strategy.

### 2. 🧠 Intelligent DocType Auto-Detection
- Employs a weighted scoring engine that analyzes header column names, primary key fields, sheet names, and filename word intersections to infer target DocType schemas automatically.

### 3. 🔑 Duplicate Control (`ignore_duplicates`)
- Option to **Ignore Duplicates (Primary Key)** before database insertion.
- Checks primary key (`name`) and title field existence (`frappe.db.exists`) and handles database unique key constraint exceptions gracefully without failing the import batch.

### 4. ⚡ Memory-Efficient High-Speed Batch Streaming
- Streams rows line-by-line without loading entire files into memory.
- Inserts records in configurable chunk sizes (default `5,000` rows) with explicit memory cleanup (`gc.collect()`).

### 5. 🎯 Dynamic UI & Real-Time Progress Tracking
- Real-time socket event emissions (`smart_import_progress`) update progress percentages, inserted record counts, and failure counts live.
- Context-aware action buttons: `"🚀 Start Import Now"` and `"🔄 Re-Analyze Files"` are automatically hidden/cleared when an import completes or is actively processing.

### 6. 📊 Detailed Failure Logging & Log Export
- Generates downloadable Excel logs (`Failed_Rows_<DocName>.xlsx`) detailing exact row numbers and error tracebacks for failed records.

---

## ⚙️ Advanced Configuration Settings

| Setting Field | Field Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `batch_size` | Int | `5000` | Number of rows processed per database commit chunk. |
| `import_type` | Select | `Insert New Records` | Import mode option (`Insert New Records`, `Update Existing Records`, `Insert and Update`). |
| `ignore_duplicates` | Check | `0` | Skip existing records matching primary key or title without throwing batch errors. |
| `ignore_link_errors` | Check | `0` | Ignore missing link validation errors during document insertion. |
| `stop_on_error` | Check | `0` | Immediately halt the entire import workflow on the first encountered error. |
| `auto_detect_doctype` | Check | `1` | Automatically match Excel/CSV headers & filenames to Frappe DocType schemas. |
| `skip_empty_rows` | Check | `1` | Automatically skip empty or blank spreadsheet rows. |
| `clean_whitespace` | Check | `1` | Automatically trim leading and trailing whitespace from string values. |

---

## 🛠️ API & Whitelisted Endpoints

```python
# Analyze attached files and build execution DAG
frappe.call({
    "method": "smart_data_import.smart_import_engine.analyze_smart_import",
    "args": {"doc_name": "SDI-2026-00001"}
})

# Start background asynchronous batch import
frappe.call({
    "method": "smart_data_import.smart_import_engine.start_smart_import",
    "args": {"doc_name": "SDI-2026-00001"}
})
```

---

## 🧪 Testing & Verification

Run the test suite using standard Frappe test runner:

```bash
bench --site <site-name> run-tests --app smart_data_import
```

Or via direct unittest module:

```bash
python -m unittest apps/ai_mcp/ai_mcp/tests/test_smart_data_import.py
```

---

## 📜 License

MIT License - ERPNext AI Team & Contributors.

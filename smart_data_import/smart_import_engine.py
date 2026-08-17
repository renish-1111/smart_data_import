# Copyright (c) 2026, ERPNext AI Team and contributors
# For license information, please see license.txt

import csv
import gc
import io
import json
import os
import re
import time
from collections import defaultdict, deque

import openpyxl
from openpyxl.styles import Font, PatternFill

import frappe
from frappe import _
from frappe.utils import cint, flt, get_files_path

EXCEL_EXTENSIONS = (".xlsx", ".xlsm", ".xltx", ".xltm")
CSV_EXTENSIONS = (".csv",)
SUPPORTED_EXTENSIONS = EXCEL_EXTENSIONS + CSV_EXTENSIONS

# Layout-only fieldtypes never hold data and must not take part in column mapping.
# How many problem rows are kept on the document for the UI. The full list always
# lives in the attached Excel log.
ERROR_ROWS_IN_UI = 500

LAYOUT_FIELDTYPES = {
	"Section Break",
	"Column Break",
	"Tab Break",
	"HTML",
	"Button",
	"Fold",
	"Heading",
	"Image",
}


def normalize_key(value):
	"""Normalizes a header/label/fieldname so `Customer Name`, `customer_name`
	and `CUSTOMER  NAME` all collapse to the same lookup key."""
	text = str(value or "").strip().lower().replace("_", " ").replace("-", " ")
	return re.sub(r"\s+", " ", text)


def resolve_file_path(file_url):
	"""Resolves an attached file URL to an absolute path on disk.

	Returns (path, error_message). Exactly one of the two is set.
	"""
	if not file_url:
		return None, _("No file attached in this row.")

	extension = os.path.splitext(file_url.split("?")[0])[1].lower()
	if extension == ".xls":
		return None, _("Legacy .xls files are not supported. Please re-save the file as .xlsx and attach it again.")
	if extension not in SUPPORTED_EXTENSIONS:
		return None, _("Unsupported file type {0}. Attach an Excel (.xlsx, .xlsm) or .csv file.").format(
			extension or file_url
		)

	file_path = None
	file_id = frappe.db.get_value("File", {"file_url": file_url}, "name")
	if file_id:
		file_path = frappe.get_doc("File", file_id).get_full_path()
	else:
		site_relative = file_url.lstrip("/")
		file_path = frappe.get_site_path("public", site_relative)
		if not os.path.exists(file_path):
			file_path = frappe.get_site_path(site_relative)

	if not file_path or not os.path.exists(file_path):
		return None, _("File could not be found on disk. Please remove the row and upload the file again.")

	return file_path, None


def build_field_lookup(meta):
	"""Builds {normalized header -> fieldname} for a DocType, covering both
	field names and human readable labels."""
	lookup = {}
	for field in meta.fields:
		if field.fieldtype in LAYOUT_FIELDTYPES:
			continue
		lookup[normalize_key(field.fieldname)] = field.fieldname
		if field.label:
			lookup.setdefault(normalize_key(field.label), field.fieldname)

	# `ID` is the label Frappe itself uses for the primary key in exports.
	lookup["name"] = "name"
	lookup["id"] = "name"
	return lookup


def map_columns(meta, headers, column_overrides=None):
	"""Maps spreadsheet headers to DocType fieldnames.

	Returns a dict with:
	        col_map  - positional list (fieldname or None) used while streaming rows
	        mapped   - {header: fieldname}
	        unmapped - [header, ...] columns that will be ignored
	"""
	lookup = build_field_lookup(meta)
	overrides = {normalize_key(k): v for k, v in (column_overrides or {}).items()}

	col_map = []
	mapped = {}
	unmapped = []

	for header in headers:
		header_text = str(header).strip() if header is not None else ""
		key = normalize_key(header_text)
		fieldname = overrides.get(key) or lookup.get(key)

		col_map.append(fieldname)
		if fieldname:
			mapped[header_text] = fieldname
		elif header_text:
			unmapped.append(header_text)

	return {"col_map": col_map, "mapped": mapped, "unmapped": unmapped}


def missing_mandatory_fields(meta, mapped_fieldnames, defaults=None):
	"""Returns labels of mandatory fields that are neither in the file nor in the
	configured default values, so the user is warned *before* the import runs."""
	provided = set(mapped_fieldnames) | set((defaults or {}).keys())
	missing = []
	for field in meta.fields:
		if not field.reqd or field.fieldtype in LAYOUT_FIELDTYPES:
			continue
		if field.default:
			continue
		if field.fieldname in provided:
			continue
		missing.append(field.label or field.fieldname)
	return missing


def sniff_csv_dialect(file_path):
	"""Detects the delimiter so semicolon/tab separated exports work out of the box."""
	try:
		with open(file_path, encoding="utf-8-sig", errors="ignore") as f:
			sample = f.read(8192)
		return csv.Sniffer().sniff(sample, delimiters=",;\t|")
	except Exception:
		return csv.excel


class SmartImportEngine:
	"""
	High-performance batch data import engine.
	Handles automated dependency graph resolution (DAG), self-referential hierarchy
	sorting, memory-efficient streaming and chunked batch insertion.
	"""

	def __init__(self, import_doc):
		if isinstance(import_doc, str):
			self.doc = frappe.get_doc("Smart Data Import", import_doc)
		else:
			self.doc = import_doc

		self.batch_size = max(500, cint(self.doc.batch_size) or 5000)
		self.skip_empty_rows = bool(self.doc.skip_empty_rows)
		self.clean_whitespace = bool(self.doc.clean_whitespace)
		self.ignore_link_errors = bool(self.doc.ignore_link_errors)
		self.ignore_mandatory = bool(getattr(self.doc, "ignore_mandatory_errors", False))
		self.stop_on_error = bool(self.doc.stop_on_error)
		self.ignore_duplicates = bool(getattr(self.doc, "ignore_duplicates", False))
		self.import_type = self.doc.import_type or "Insert New Records"
		self.rules = self._parse_filter_rules()
		self._dt_index = None
		self._manifest_buffer = []

	# -------------------------------------------------------------------------
	# 0. CONFIGURATION HELPERS
	# -------------------------------------------------------------------------

	def _parse_filter_rules(self):
		"""Parses the optional per-DocType mapping/defaults configuration:

		{
		    "Customer": {
		        "column_map": {"Cust Name": "customer_name"},
		        "defaults": {"customer_type": "Company"}
		    }
		}
		"""
		if not self.doc.filter_rules_json:
			return {}
		try:
			rules = json.loads(self.doc.filter_rules_json)
			return rules if isinstance(rules, dict) else {}
		except Exception:
			return {}

	def rules_for(self, target_doctype):
		rules = self.rules.get(target_doctype) or {}
		if not isinstance(rules, dict):
			return {}, {}
		column_map = rules.get("column_map") or {}
		defaults = rules.get("defaults") or {}
		return (
			column_map if isinstance(column_map, dict) else {},
			defaults if isinstance(defaults, dict) else {},
		)

	def supports(self, fieldname):
		"""True when the site's schema actually has this field.

		Code can reach a site before `bench migrate` does; when that happens the newer
		features degrade quietly instead of crashing a running import.
		"""
		return bool(self.doc.meta.has_field(fieldname))

	def supported_updates(self, values):
		"""Filters an update dict down to columns that exist on this site."""
		return {k: v for k, v in values.items() if self.supports(k)}

	@property
	def processed_records(self):
		return (
			cint(self.doc.imported_records)
			+ cint(self.doc.failed_records)
			+ cint(getattr(self.doc, "skipped_records", 0))
		)

	# -------------------------------------------------------------------------
	# 1. FILE ANALYSIS & DOCTYPE AUTO-DETECTION
	# -------------------------------------------------------------------------

	def analyze_files_and_build_graph(self):
		"""
		Reads attached files, infers target DocTypes, counts rows, validates the
		column mapping and builds the DAG used for dependency resolution.
		"""
		self.doc.status = "Analyzing"
		self.doc.save(ignore_permissions=True)
		frappe.db.commit()

		detected_doctypes = set()

		for row in self.doc.files:
			row.mapping_summary = ""
			row.error_log = ""

			file_path, error = resolve_file_path(row.file)
			if error:
				row.status = "Failed"
				row.error_log = error
				row.total_rows = 0
				continue

			try:
				headers, row_count, sheet_name = self.read_file_header_and_count(file_path, row.sheet_name)
			except Exception as e:
				row.status = "Failed"
				row.error_log = _("Could not read file: {0}").format(str(e))
				row.total_rows = 0
				continue

			row.total_rows = row_count
			if sheet_name:
				row.sheet_name = sheet_name

			# Auto-detection only fills blanks; a manually chosen DocType is never overwritten.
			if not row.doctype_name and self.doc.auto_detect_doctype:
				detected, note = self.detect_target_doctype(row.file, row.sheet_name, headers)
				if detected:
					row.doctype_name = detected
					row.mapping_summary = note

			if not row.doctype_name:
				row.status = "Failed"
				row.error_log = _(
					"Could not detect the target DocType. Please pick it manually in the 'Target DocType' column."
				)
				continue

			if not frappe.db.exists("DocType", row.doctype_name):
				row.status = "Failed"
				row.error_log = _("DocType {0} does not exist.").format(row.doctype_name)
				continue

			row.status = "Analyzed"
			detected_doctypes.add(row.doctype_name)

			if not row_count:
				row.status = "Failed"
				row.error_log = _("No data rows found. Row 1 must contain the column headers.")
				continue

			row.mapping_summary = self._describe_mapping(row.doctype_name, headers, row.mapping_summary)

		# Build dependency graph & topological tiers
		self.doc.dependencies = []
		total_records = 0
		if detected_doctypes:
			for dt_info in self.build_topological_dependency_tiers(detected_doctypes):
				total_dt_rows = sum(
					cint(r.total_rows) for r in self.doc.files if r.doctype_name == dt_info["doctype"]
				)
				total_records += total_dt_rows
				self.doc.append(
					"dependencies",
					{
						"execution_tier": dt_info["tier"],
						"doctype_name": dt_info["doctype"],
						"depends_on_doctypes": ", ".join(dt_info["depends_on"]),
						"has_inner_dependency": 1 if dt_info["inner_ref_field"] else 0,
						"self_reference_field": dt_info["inner_ref_field"] or "",
						"status": "Pending",
						"total_count": total_dt_rows,
						"processed_count": 0,
					},
				)

		self.doc.total_records = total_records
		self.doc.status = "Ready" if total_records else "Pending"
		self.doc.save(ignore_permissions=True)
		frappe.db.commit()
		return True

	def _describe_mapping(self, target_doctype, headers, prefix=""):
		"""Builds the short human readable mapping report shown on each file row."""
		meta = frappe.get_meta(target_doctype)
		column_map, defaults = self.rules_for(target_doctype)
		mapping = map_columns(meta, headers, column_map)
		missing = missing_mandatory_fields(meta, mapping["mapped"].values(), defaults)

		notes = [prefix] if prefix else []
		notes.append(_("{0} of {1} columns mapped.").format(len(mapping["mapped"]), len(mapping["mapped"]) + len(mapping["unmapped"])))
		if mapping["unmapped"]:
			notes.append(_("Ignored columns: {0}").format(", ".join(mapping["unmapped"][:10])))
		if missing:
			notes.append(_("⚠ Mandatory fields not in file: {0}").format(", ".join(missing[:10])))
		return " ".join(notes)

	def read_file_header_and_count(self, file_path, sheet_name=None):
		"""
		Reads headers and counts data rows without loading the whole file into RAM.
		Empty header cells keep their position so column alignment is preserved.
		"""
		headers = []
		count = 0
		selected_sheet = sheet_name

		if file_path.lower().endswith(EXCEL_EXTENSIONS):
			wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
			try:
				sheet_names = wb.sheetnames
				selected_sheet = sheet_name if (sheet_name and sheet_name in sheet_names) else sheet_names[0]
				ws = wb[selected_sheet]
				for idx, row in enumerate(ws.iter_rows(values_only=True)):
					if idx == 0:
						headers = self._clean_header_row(row)
					elif any(cell is not None and str(cell).strip() != "" for cell in row):
						count += 1
			finally:
				wb.close()
		elif file_path.lower().endswith(CSV_EXTENSIONS):
			dialect = sniff_csv_dialect(file_path)
			with open(file_path, encoding="utf-8-sig", errors="ignore", newline="") as f:
				for idx, row in enumerate(csv.reader(f, dialect)):
					if idx == 0:
						headers = self._clean_header_row(row)
					elif any(cell is not None and str(cell).strip() != "" for cell in row):
						count += 1

		return headers, count, selected_sheet

	@staticmethod
	def _clean_header_row(row):
		"""Keeps positional alignment (blank cells become ''), drops trailing blanks."""
		headers = ["" if cell is None else str(cell).strip() for cell in row]
		while headers and headers[-1] == "":
			headers.pop()
		return headers

	def _get_doctype_index(self):
		"""Builds a single in-memory index of {doctype: field keys} using two queries
		instead of calling frappe.get_meta() for every DocType on the site."""
		if self._dt_index is not None:
			return self._dt_index

		index = {}
		for dt in frappe.get_all("DocType", filters={"istable": 0, "issingle": 0}, fields=["name", "title_field"]):
			index[dt.name] = {"title_field": dt.title_field, "keys": set()}

		DocField = frappe.qb.DocType("DocField")
		rows = frappe.qb.from_(DocField).select(DocField.parent, DocField.fieldname, DocField.label).run(as_dict=True)

		CustomField = frappe.qb.DocType("Custom Field")
		rows += (
			frappe.qb.from_(CustomField)
			.select(CustomField.dt.as_("parent"), CustomField.fieldname, CustomField.label)
			.run(as_dict=True)
		)

		for row in rows:
			entry = index.get(row.parent)
			if not entry:
				continue
			if row.fieldname:
				entry["keys"].add(normalize_key(row.fieldname))
			if row.label:
				entry["keys"].add(normalize_key(row.label))

		self._dt_index = index
		return index

	def detect_target_doctype(self, filename, sheet_name, headers):
		"""
		Infers the target DocType with a weighted score over header fields,
		primary-key naming, sheet name and filename word overlap.

		Returns (doctype, note) where note explains the confidence to the user.
		"""
		index = self._get_doctype_index()
		scores = defaultdict(int)

		header_set = {normalize_key(h) for h in headers if h}
		header_set.discard("")

		# Step 1: header/label matching (highest confidence signal)
		if header_set:
			for dt, entry in index.items():
				matches = len(header_set & entry["keys"])
				if normalize_key(dt) + " name" in header_set:
					matches += 3
				if entry["title_field"] and normalize_key(entry["title_field"]) in header_set:
					matches += 1
				if matches >= 2:
					scores[dt] += matches * 10

		# Step 2: sheet name matching
		if sheet_name:
			clean_sheet = normalize_key(sheet_name)
			for dt in index:
				if normalize_key(dt) == clean_sheet:
					scores[dt] += 40

		# Step 3: filename matching & word overlap
		base_name = os.path.basename(str(filename or "")).rsplit(".", 1)[0]
		clean_base = normalize_key(base_name)
		base_words = set(clean_base.split())
		for dt in index:
			dt_clean = normalize_key(dt)
			if dt_clean == clean_base:
				scores[dt] += 50
			else:
				common = set(dt_clean.split()) & base_words
				if common:
					scores[dt] += len(common) * 5

		if not scores:
			return None, ""

		ranked = sorted(scores.items(), key=lambda x: (-x[1], x[0]))
		best, best_score = ranked[0]
		note = _("Auto-detected {0}.").format(best)
		if len(ranked) > 1 and ranked[1][1] >= best_score * 0.9:
			note = _("Auto-detected {0} (close match with {1} — please confirm).").format(best, ranked[1][0])
		return best, note

	# -------------------------------------------------------------------------
	# 2. TOPOLOGICAL SORT FOR INTER & INTRA DOCTYPE DEPENDENCIES
	# -------------------------------------------------------------------------

	def build_topological_dependency_tiers(self, doctypes_set):
		"""
		Builds a DAG for the target DocTypes and runs Kahn's topological sort to
		determine execution tiers. Primary document Link fields take priority over
		child table links so cycles are avoided.
		"""
		doctypes = sorted(doctypes_set)
		graph = defaultdict(set)  # u -> nodes that depend on u
		in_degree = {dt: 0 for dt in doctypes}
		depends_on_map = defaultdict(set)
		inner_ref_map = {}

		for dt in doctypes:
			meta = frappe.get_meta(dt)
			inner_ref = None
			primary_link_doctypes = set()

			for f in meta.fields:
				if f.fieldtype == "Link" and f.options:
					if f.options == dt:
						inner_ref = f.fieldname
					elif f.options in doctypes_set:
						primary_link_doctypes.add(f.options)

			inner_ref_map[dt] = inner_ref
			depends_on_map[dt] = primary_link_doctypes

			for target_dt in primary_link_doctypes:
				if dt not in graph[target_dt]:
					graph[target_dt].add(dt)
					in_degree[dt] += 1

		queue = deque([dt for dt in doctypes if in_degree[dt] == 0])
		result_tiers = []
		current_tier = 0
		visited_count = 0

		while queue:
			for _i in range(len(queue)):
				u = queue.popleft()
				visited_count += 1
				result_tiers.append(
					{
						"doctype": u,
						"tier": current_tier,
						"depends_on": sorted(depends_on_map[u]),
						"inner_ref_field": inner_ref_map[u],
					}
				)
				for v in graph[u]:
					in_degree[v] -= 1
					if in_degree[v] == 0:
						queue.append(v)
			current_tier += 1

		# Cyclic leftovers go into a final tier so they are still imported.
		if visited_count < len(doctypes):
			for u in [dt for dt in doctypes if in_degree[dt] > 0]:
				result_tiers.append(
					{
						"doctype": u,
						"tier": current_tier,
						"depends_on": sorted(depends_on_map[u]),
						"inner_ref_field": inner_ref_map[u],
					}
				)

		return result_tiers

	# -------------------------------------------------------------------------
	# 3. HIGH-SPEED BATCH IMPORT EXECUTION ENGINE
	# -------------------------------------------------------------------------

	def execute_import(self):
		"""
		Executes the import tier by tier, streaming rows, flushing chunks and
		publishing real-time progress.
		"""
		start_time = time.time()
		self.doc.status = "Processing"
		self.doc.imported_records = 0
		self.doc.failed_records = 0
		self.doc.skipped_records = 0
		self.doc.progress_percent = 0
		self.doc.error_log = ""
		self.doc.save(ignore_permissions=True)
		frappe.db.commit()

		sorted_deps = sorted(self.doc.dependencies, key=lambda d: cint(d.execution_tier))
		row_log = []

		for dep in sorted_deps:
			dep.status = "Processing"
			self.doc.save(ignore_permissions=True)
			frappe.db.commit()

			target_dt = dep.doctype_name
			self.publish_progress(_("Tier {0}: importing {1}...").format(dep.execution_tier, target_dt))

			matching_files = [f for f in self.doc.files if f.doctype_name == target_dt]
			for file_row in matching_files:
				file_path, error = resolve_file_path(file_row.file)
				if error:
					file_row.status = "Failed"
					file_row.error_log = error
					row_log.append({"row": 0, "doctype": target_dt, "type": "Failed", "reason": error})
					continue

				file_row.status = "Importing"
				# Counters are advanced chunk-by-chunk inside _update_progress_counts,
				# so the totals returned here are used for reporting only.
				count, failed, skipped, errors = self._stream_and_batch_insert(
					file_path=file_path,
					sheet_name=file_row.sheet_name,
					target_doctype=target_dt,
					self_ref_field=dep.self_reference_field if dep.has_inner_dependency else None,
					dep_row=dep,
				)
				row_log.extend(errors)

				file_row.status = "Failed" if (failed and not count) else "Completed"
				file_row.error_log = _("{0} imported, {1} failed, {2} skipped.").format(count, failed, skipped)

				if failed and self.stop_on_error:
					dep.status = "Failed"
					self._finalize(start_time, row_log, "Failed")
					return False

			dep.status = "Completed" if cint(dep.processed_count) else "Pending"
			self.doc.save(ignore_permissions=True)
			frappe.db.commit()

		failed_only = [e for e in row_log if e.get("type") != "Skipped"]
		if failed_only:
			status = "Partial Success" if cint(self.doc.imported_records) else "Failed"
		else:
			status = "Completed"

		self._finalize(start_time, row_log, status)
		return True

	def _finalize(self, start_time, row_log, status):
		self._flush_manifest()
		self._attach_manifest()
		self._store_error_rows(row_log)
		self.doc.execution_time_seconds = flt(time.time() - start_time, 2)
		self.doc.status = status
		if row_log:
			self._generate_failed_rows_excel(row_log)
			self.doc.error_log = "\n".join(
				f"Row {entry['row']} [{entry.get('type', 'Failed')}]: {entry['reason']}" for entry in row_log[:200]
			)
			if len(row_log) > 200:
				self.doc.error_log += "\n" + _("... {0} more entries, see the downloadable log.").format(
					len(row_log) - 200
				)
		self.doc.progress_percent = 100
		self.doc.save(ignore_permissions=True)
		frappe.db.commit()
		self.publish_progress(
			_("Import finished: {0} imported, {1} failed, {2} skipped.").format(
				self.doc.imported_records, self.doc.failed_records, self.doc.skipped_records
			),
			100,
		)

	def _stream_and_batch_insert(self, file_path, sheet_name, target_doctype, self_ref_field, dep_row):
		"""
		Streams rows in configurable batch sizes and handles self-referential
		hierarchies with a two-pass strategy.
		"""
		meta = frappe.get_meta(target_doctype)
		column_map, defaults = self.rules_for(target_doctype)

		inserted_count = 0
		failed_count = 0
		skipped_count = 0
		row_log = []

		row_generator = self._get_row_generator(file_path, sheet_name)
		headers = next(row_generator, None)
		if not headers:
			return 0, 0, 0, []

		mapping = map_columns(meta, self._clean_header_row(headers), column_map)
		col_map = mapping["col_map"]
		if not mapping["mapped"]:
			reason = _("None of the columns in this file match fields of {0}.").format(target_doctype)
			self._add_counts(dep_row, 0, 1, 0)
			return 0, 1, 0, [{"row": 1, "doctype": target_doctype, "type": "Failed", "reason": reason}]

		title_field = self._get_title_field(meta, target_doctype)

		batch_records = []
		second_pass_updates = []
		row_idx = 1  # header occupies row 1

		for row_values in row_generator:
			row_idx += 1
			row_dict = dict(defaults)
			for idx, val in enumerate(row_values):
				fieldname = col_map[idx] if idx < len(col_map) else None
				if not fieldname or val is None:
					continue
				value = val.strip() if (self.clean_whitespace and isinstance(val, str)) else val
				if value != "":
					row_dict[fieldname] = value

			if self.skip_empty_rows and not any(k not in defaults for k in row_dict):
				continue

			row_dict = self._resolve_row_link_fields(target_doctype, row_dict, meta)

			# Pass 1 drops the self reference, pass 2 sets it once all rows exist.
			if self_ref_field and row_dict.get(self_ref_field):
				parent_val = row_dict.pop(self_ref_field)
				child_key = row_dict.get("name") or (row_dict.get(title_field) if title_field else None)
				if child_key:
					second_pass_updates.append((child_key, self_ref_field, parent_val))

			batch_records.append((row_idx, row_dict))

			if len(batch_records) >= self.batch_size:
				c, f, s, errs = self._flush_batch_to_db(target_doctype, batch_records, meta, title_field)
				inserted_count += c
				failed_count += f
				skipped_count += s
				row_log.extend(errs)
				batch_records = []
				self._update_progress_counts(dep_row, c, f, s)
				gc.collect()
				if f and self.stop_on_error:
					return inserted_count, failed_count, skipped_count, row_log

		if batch_records:
			c, f, s, errs = self._flush_batch_to_db(target_doctype, batch_records, meta, title_field)
			inserted_count += c
			failed_count += f
			skipped_count += s
			row_log.extend(errs)
			self._update_progress_counts(dep_row, c, f, s)

		if second_pass_updates:
			self.publish_progress(_("Linking hierarchy parents for {0}...").format(target_doctype))
			row_log.extend(self._apply_second_pass(target_doctype, meta, title_field, second_pass_updates))

		return inserted_count, failed_count, skipped_count, row_log

	@staticmethod
	def _get_title_field(meta, target_doctype):
		if meta.title_field:
			return meta.title_field
		for candidate in (
			target_doctype.lower().replace(" ", "_") + "_name",
			"title",
			"subject",
			"full_name",
		):
			if meta.has_field(candidate):
				return candidate
		return None

	def _apply_second_pass(self, target_doctype, meta, title_field, updates):
		"""Sets self-referential parent links now that every row exists.

		Tree DocTypes (Item Group, Customer Group, Account, Task, ...) are saved
		through the document so the nested set (lft/rgt) stays consistent — writing
		the column directly would leave the tree corrupted. Flat DocTypes take the
		fast direct-write path.
		"""
		is_tree = bool(getattr(meta, "is_tree", 0)) or (meta.has_field("lft") and meta.has_field("rgt"))
		problems = []

		for child_key, fieldname, parent_val in updates:
			try:
				child_name = child_key
				if not frappe.db.exists(target_doctype, child_name) and title_field:
					child_name = frappe.db.get_value(target_doctype, {title_field: child_key}, "name")

				parent_name = parent_val
				if not frappe.db.exists(target_doctype, parent_name) and title_field:
					parent_name = frappe.db.get_value(target_doctype, {title_field: parent_val}, "name")

				if not child_name or not parent_name:
					problems.append(
						{
							"row": 0,
							"doctype": target_doctype,
							"type": "Hierarchy",
							"reason": _("Could not link {0} to parent {1} — record not found.").format(
								child_key, parent_val
							),
						}
					)
					continue

				if is_tree:
					child = frappe.get_doc(target_doctype, child_name)
					child.set(fieldname, parent_name)
					child.flags.ignore_permissions = True
					child.save(ignore_permissions=True)
				else:
					frappe.db.set_value(
						target_doctype, child_name, fieldname, parent_name, update_modified=False
					)
			except Exception as e:
				problems.append(
					{
						"row": 0,
						"doctype": target_doctype,
						"type": "Hierarchy",
						"reason": _("Parent link {0} -> {1} failed: {2}").format(
							child_key, parent_val, self._clean_error(e)
						),
					}
				)

		frappe.db.commit()
		return problems

	def _resolve_row_link_fields(self, target_doctype, row_dict, meta):
		"""
		Resolves human readable Link titles to primary keys for autonamed DocTypes,
		e.g. Task.project = "Kintech AI Engine" -> "PROJ-0001".
		"""
		resolved = dict(row_dict)
		for f in meta.fields:
			if f.fieldtype != "Link" or not f.options or f.fieldname not in resolved:
				continue
			val = resolved[f.fieldname]
			if not val or not isinstance(val, str):
				continue
			link_dt = f.options
			if frappe.db.exists(link_dt, val):
				continue

			link_meta = frappe.get_meta(link_dt)
			search_fields = []
			if link_meta.title_field:
				search_fields.append(link_meta.title_field)
			search_fields.append(link_dt.lower().replace(" ", "_") + "_name")
			search_fields.extend(["title", "subject", "item_name", "full_name"])

			for sf in search_fields:
				if link_meta.has_field(sf):
					res = frappe.db.get_value(link_dt, {sf: val}, "name")
					if res:
						resolved[f.fieldname] = res
						break
		return resolved

	def _find_existing_name(self, target_doctype, meta, row_dict, title_field):
		"""Locates an existing record for duplicate detection / update modes."""
		name = row_dict.get("name")
		if name and frappe.db.exists(target_doctype, name):
			return name

		for f in meta.fields:
			if f.unique and row_dict.get(f.fieldname):
				existing = frappe.db.get_value(target_doctype, {f.fieldname: row_dict[f.fieldname]}, "name")
				if existing:
					return existing

		if title_field and row_dict.get(title_field):
			return frappe.db.get_value(target_doctype, {title_field: row_dict[title_field]}, "name")

		return None

	def _flush_batch_to_db(self, target_doctype, batch_records, meta, title_field):
		"""Writes one chunk of records, honouring the selected import mode."""
		c_success = 0
		c_failed = 0
		c_skipped = 0
		row_log = []

		needs_lookup = self.ignore_duplicates or self.import_type != "Insert New Records"

		for row_idx, row_dict in batch_records:
			try:
				existing = (
					self._find_existing_name(target_doctype, meta, row_dict, title_field)
					if needs_lookup
					else None
				)

				if existing and self.import_type in ("Update Existing Records", "Insert and Update"):
					self._update_existing(target_doctype, existing, row_dict)
					self._record_in_manifest("U", target_doctype, existing)
					c_success += 1
					continue

				if existing and self.ignore_duplicates:
					c_skipped += 1
					row_log.append(
						{
							"row": row_idx,
							"doctype": target_doctype,
							"type": "Skipped",
							"reason": _("Already exists as {0}").format(existing),
						}
					)
					continue

				if not existing and self.import_type == "Update Existing Records":
					c_skipped += 1
					row_log.append(
						{
							"row": row_idx,
							"doctype": target_doctype,
							"type": "Skipped",
							"reason": _("No matching record found to update."),
						}
					)
					continue

				doc = frappe.new_doc(target_doctype)
				doc.update(row_dict)
				doc.flags.ignore_permissions = True
				doc.insert(
					ignore_permissions=True,
					ignore_mandatory=self.ignore_mandatory,
					ignore_links=self.ignore_link_errors,
				)
				self._record_in_manifest("I", target_doctype, doc.name)
				c_success += 1

			except Exception as e:
				is_duplicate = isinstance(e, frappe.DuplicateEntryError) or "Duplicate" in type(e).__name__
				if is_duplicate and self.ignore_duplicates:
					c_skipped += 1
					row_log.append(
						{
							"row": row_idx,
							"doctype": target_doctype,
							"type": "Skipped",
							"reason": _("Duplicate entry skipped."),
						}
					)
				else:
					c_failed += 1
					row_log.append(
						{
							"row": row_idx,
							"doctype": target_doctype,
							"type": "Failed",
							"reason": self._clean_error(e),
						}
					)
					if self.stop_on_error:
						self._flush_manifest()
						frappe.db.commit()
						return c_success, c_failed, c_skipped, row_log

		self._flush_manifest()
		frappe.db.commit()
		return c_success, c_failed, c_skipped, row_log

	def _update_existing(self, target_doctype, name, row_dict):
		doc = frappe.get_doc(target_doctype, name)
		payload = {k: v for k, v in row_dict.items() if k != "name"}
		doc.update(payload)
		doc.flags.ignore_permissions = True
		doc.save(ignore_permissions=True)

	@staticmethod
	def _clean_error(exception):
		"""Turns Frappe's HTML-flavoured validation messages into a readable line."""
		message = str(exception) or type(exception).__name__
		message = re.sub(r"<[^>]+>", " ", message)
		message = re.sub(r"\s+", " ", message).strip()
		return message[:500]

	def _get_row_generator(self, file_path, sheet_name):
		"""Yields rows one at a time for Excel or CSV files."""
		if file_path.lower().endswith(EXCEL_EXTENSIONS):
			wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
			try:
				ws = wb[sheet_name] if (sheet_name and sheet_name in wb.sheetnames) else wb.active
				yield from ws.iter_rows(values_only=True)
			finally:
				wb.close()
		elif file_path.lower().endswith(CSV_EXTENSIONS):
			dialect = sniff_csv_dialect(file_path)
			with open(file_path, encoding="utf-8-sig", errors="ignore", newline="") as f:
				yield from csv.reader(f, dialect)

	def _add_counts(self, dep_row, inserted, failed, skipped):
		"""Single place where the running counters are advanced, so nothing is
		counted twice."""
		dep_row.processed_count = cint(dep_row.processed_count) + inserted + failed + skipped
		self.doc.imported_records = cint(self.doc.imported_records) + inserted
		self.doc.failed_records = cint(self.doc.failed_records) + failed
		self.doc.skipped_records = cint(self.doc.skipped_records) + skipped

	def _update_progress_counts(self, dep_row, inserted, failed, skipped):
		"""Publishes progress using lightweight writes so an open form is not
		flagged as modified while the background job runs."""
		self._add_counts(dep_row, inserted, failed, skipped)

		if cint(self.doc.total_records) > 0:
			self.doc.progress_percent = min(
				99.0, flt((self.processed_records / cint(self.doc.total_records)) * 100, 1)
			)

		frappe.db.set_value(
			"Smart Data Import",
			self.doc.name,
			self.supported_updates(
				{
					"imported_records": self.doc.imported_records,
					"failed_records": self.doc.failed_records,
					"skipped_records": self.doc.skipped_records,
					"progress_percent": self.doc.progress_percent,
				}
			),
			update_modified=False,
		)
		frappe.db.set_value(
			"Smart Data Import Dependency",
			dep_row.name,
			"processed_count",
			dep_row.processed_count,
			update_modified=False,
		)
		frappe.db.commit()
		self.publish_progress(
			_("{0}: {1} of {2} rows processed...").format(
				dep_row.doctype_name, dep_row.processed_count, dep_row.total_count
			)
		)

	def publish_progress(self, message, progress_pct=None):
		# A realtime/redis hiccup must never abort a long running import.
		try:
			self._publish(message, progress_pct)
		except Exception:
			pass

	def _publish(self, message, progress_pct=None):
		frappe.publish_realtime(
			"smart_import_progress",
			{
				"doc_name": self.doc.name,
				"message": message,
				"progress": progress_pct if progress_pct is not None else self.doc.progress_percent,
				"imported": cint(self.doc.imported_records),
				"failed": cint(self.doc.failed_records),
				"skipped": cint(getattr(self.doc, "skipped_records", 0)),
				"rolled_back": cint(getattr(self.doc, "rolled_back_records", 0)),
				"status": self.doc.status,
			},
			user=self.doc.owner,
			after_commit=False,
		)

	def _store_error_rows(self, row_log):
		"""Puts the problem rows on the document itself, so the reasons are readable
		in the form without downloading anything.

		Only the first ERROR_ROWS_IN_UI entries are stored — the complete list stays
		in the attached Excel log, which is what large imports need anyway.
		"""
		if not self.supports("errors"):
			return

		self.doc.errors = []
		for entry in row_log[:ERROR_ROWS_IN_UI]:
			self.doc.append(
				"errors",
				{
					"row_index": cint(entry.get("row")),
					"error_type": entry.get("type") or "Failed",
					"doctype_name": entry.get("doctype"),
					"reason": entry.get("reason"),
				},
			)

	def _generate_failed_rows_excel(self, row_log):
		"""Attaches a private Excel log listing every failed and skipped row."""
		wb = openpyxl.Workbook()
		ws = wb.active
		ws.title = "Import Log"
		ws.append(["Row Index", "Target DocType", "Type", "Reason"])
		for entry in row_log:
			ws.append(
				[entry.get("row"), entry.get("doctype"), entry.get("type", "Failed"), entry.get("reason")]
			)

		stream = io.BytesIO()
		wb.save(stream)
		wb.close()

		self._remove_previous_log()
		from frappe.utils.file_manager import save_file

		file_doc = save_file(
			f"Import_Log_{self.doc.name}.xlsx",
			stream.getvalue(),
			"Smart Data Import",
			self.doc.name,
			is_private=1,
		)
		self.doc.error_file = file_doc.file_url

	def _remove_previous_log(self):
		for name in frappe.get_all(
			"File",
			filters={
				"attached_to_doctype": "Smart Data Import",
				"attached_to_name": self.doc.name,
				"file_name": ("like", "Import_Log_%"),
			},
			pluck="name",
		):
			try:
				frappe.delete_doc("File", name, force=True, ignore_permissions=True)
			except Exception:
				pass

	# -------------------------------------------------------------------------
	# 4. ROLLBACK MANIFEST
	# -------------------------------------------------------------------------

	def _record_in_manifest(self, action, target_doctype, name):
		"""Remembers one written record so the import can be undone later.

		`action` is "I" for an inserted record (deletable on rollback) or "U" for an
		updated one (its previous values are not recoverable).
		"""
		self._manifest_buffer.append((action, target_doctype, name))

	def _flush_manifest(self):
		"""Appends the buffered entries to the manifest file on disk.

		Written incrementally and never held in memory, so a manifest for millions of
		rows costs nothing and survives a crash mid-import.
		"""
		if not self._manifest_buffer:
			return
		path = manifest_path(self.doc.name)
		os.makedirs(os.path.dirname(path), exist_ok=True)
		with open(path, "a", encoding="utf-8", newline="") as f:
			csv.writer(f).writerows(self._manifest_buffer)
		self._manifest_buffer = []

	def _attach_manifest(self):
		"""Publishes the manifest as a private attachment so it is visible and
		downloadable, and stores its URL on the document."""
		if not self.supports("rollback_file"):
			return

		path = manifest_path(self.doc.name)
		if not os.path.exists(path):
			return

		file_url = f"/private/files/{os.path.basename(path)}"
		if getattr(self.doc, "rollback_file", None) == file_url and frappe.db.exists(
			"File", {"file_url": file_url}
		):
			return

		if not frappe.db.exists("File", {"file_url": file_url}):
			frappe.get_doc(
				{
					"doctype": "File",
					"file_name": os.path.basename(path),
					"file_url": file_url,
					"is_private": 1,
					"attached_to_doctype": "Smart Data Import",
					"attached_to_name": self.doc.name,
				}
			).insert(ignore_permissions=True)
		self.doc.rollback_file = file_url

	# -------------------------------------------------------------------------
	# 5. ROLLBACK EXECUTION
	# -------------------------------------------------------------------------

	def execute_rollback(self, force=False):
		"""Deletes every record this import created, newest first.

		Newest-first is what makes it safe: child rows, later tiers and tree children
		go before the parents they point at. Records that cannot be deleted (still
		linked elsewhere) are retried in further passes and then reported.
		"""
		start_time = time.time()
		entries = read_manifest(self.doc.name)
		inserted = [e for e in entries if e[0] == "I"]
		updated_count = sum(1 for e in entries if e[0] == "U")

		self.doc.status = "Rolling Back"
		self.doc.progress_percent = 0
		self.doc.rolled_back_records = 0
		self.doc.save(ignore_permissions=True)
		frappe.db.commit()

		# Fail before deleting anything if a DocType may not be deleted from.
		for target_doctype in {e[1] for e in inserted}:
			frappe.has_permission(target_doctype, "delete", throw=True)

		total = len(inserted)
		self.publish_progress(_("Rolling back {0} records...").format(total))

		# Hierarchies link both ways (a child points at its parent, and the parent's
		# own validations refuse to go while a child exists), so no deletion order
		# alone can break the pair. Detach the self-references first.
		self._unlink_self_references(inserted)

		pending = list(reversed(inserted))
		deleted = 0
		problems = {}

		for _pass_no in range(3):
			retry = []
			for _action, target_doctype, name in pending:
				try:
					if self._delete_one(target_doctype, name, force):
						deleted += 1
						problems.pop((target_doctype, name), None)
					frappe.db.commit()
				except Exception as e:
					frappe.db.rollback()
					problems[(target_doctype, name)] = self._clean_error(e)
					retry.append((_action, target_doctype, name))

				if total and deleted % 200 == 0:
					self._publish_rollback_progress(deleted, total)

			if not retry or len(retry) == len(pending):
				pending = retry
				break
			pending = retry

		row_log = [
			{
				"row": 0,
				"doctype": target_doctype,
				"type": "Rollback",
				"reason": _("Could not delete {0}: {1}").format(name, reason),
			}
			for (target_doctype, name), reason in problems.items()
		]

		# Keep only what is still there, so a second rollback attempt retries just those.
		rewrite_manifest(self.doc.name, pending)

		self.doc.reload()
		self.doc.rolled_back_records = deleted
		self.doc.execution_time_seconds = flt(time.time() - start_time, 2)
		self.doc.status = "Rolled Back"
		self.doc.progress_percent = 100
		self.doc.error_log = ""
		self._store_error_rows(row_log)
		if row_log:
			self._generate_failed_rows_excel(row_log)
			self.doc.error_log = "\n".join(entry["reason"] for entry in row_log[:200])
		if updated_count:
			self.doc.error_log = (
				_("{0} record(s) were updated by this import — updates cannot be undone automatically.").format(
					updated_count
				)
				+ ("\n" + self.doc.error_log if self.doc.error_log else "")
			)
		if not pending:
			self._detach_manifest()
		else:
			self._attach_manifest()
		self.doc.save(ignore_permissions=True)
		frappe.db.commit()

		self.publish_progress(
			_("Rollback finished: {0} of {1} records deleted, {2} could not be deleted.").format(
				deleted, total, len(pending)
			),
			100,
		)
		return {"deleted": deleted, "total": total, "failed": len(pending), "updated": updated_count}

	def _unlink_self_references(self, inserted):
		"""Clears self-referencing Link fields (parent_task, parent_item_group, ...)
		on the records that are about to be deleted.

		Tree DocTypes are saved through the document so the nested set is updated;
		flat ones are written directly. Only records whose field is actually set are
		touched, and records outside this import are never modified — a link from
		outside must still block the delete.
		"""
		by_doctype = defaultdict(list)
		for _action, target_doctype, name in inserted:
			by_doctype[target_doctype].append(name)

		for target_doctype, names in by_doctype.items():
			meta = frappe.get_meta(target_doctype)
			self_fields = [
				f.fieldname
				for f in meta.fields
				if f.fieldtype == "Link" and f.options == target_doctype
			]
			if not self_fields:
				continue

			is_tree = bool(getattr(meta, "is_tree", 0)) or (
				meta.has_field("lft") and meta.has_field("rgt")
			)

			for chunk_start in range(0, len(names), 500):
				chunk = names[chunk_start : chunk_start + 500]
				rows = frappe.get_all(
					target_doctype,
					filters={"name": ("in", chunk)},
					fields=["name", *self_fields],
				)
				for row in rows:
					if not any(row.get(f) for f in self_fields):
						continue
					try:
						if is_tree:
							doc = frappe.get_doc(target_doctype, row["name"])
							for fieldname in self_fields:
								doc.set(fieldname, None)
							doc.flags.ignore_permissions = True
							doc.flags.ignore_links = True
							doc.save(ignore_permissions=True)
						else:
							frappe.db.set_value(
								target_doctype,
								row["name"],
								{f: None for f in self_fields if row.get(f)},
								update_modified=False,
							)
					except Exception:
						# Fall back to a direct write; the record is being deleted anyway.
						frappe.db.rollback()
						frappe.db.set_value(
							target_doctype,
							row["name"],
							{f: None for f in self_fields if row.get(f)},
							update_modified=False,
						)
				frappe.db.commit()

	def _delete_one(self, target_doctype, name, force):
		"""Deletes a single record, cancelling it first if it was submitted."""
		if not frappe.db.exists(target_doctype, name):
			return True  # already gone — nothing to undo

		if cint(frappe.db.get_value(target_doctype, name, "docstatus")) == 1:
			doc = frappe.get_doc(target_doctype, name)
			doc.flags.ignore_permissions = True
			doc.cancel()

		frappe.delete_doc(
			target_doctype,
			name,
			force=1 if force else 0,
			ignore_permissions=True,
			ignore_missing=True,
			delete_permanently=True,
		)
		return True

	def _publish_rollback_progress(self, deleted, total):
		percent = flt((deleted / total) * 100, 1) if total else 100
		frappe.db.set_value(
			"Smart Data Import",
			self.doc.name,
			self.supported_updates({"rolled_back_records": deleted, "progress_percent": percent}),
			update_modified=False,
		)
		frappe.db.commit()
		self.doc.progress_percent = percent
		self.doc.rolled_back_records = deleted
		self.publish_progress(_("Deleted {0} of {1} records...").format(deleted, total), percent)

	def _detach_manifest(self):
		"""Removes the manifest once everything it listed is gone."""
		for name in frappe.get_all(
			"File",
			filters={
				"attached_to_doctype": "Smart Data Import",
				"attached_to_name": self.doc.name,
				"file_name": ("like", "Rollback_%"),
			},
			pluck="name",
		):
			try:
				frappe.delete_doc("File", name, force=True, ignore_permissions=True)
			except Exception:
				pass
		path = manifest_path(self.doc.name)
		if os.path.exists(path):
			os.remove(path)
		self.doc.rollback_file = None


# -------------------------------------------------------------------------
# ROLLBACK MANIFEST HELPERS
# -------------------------------------------------------------------------


def manifest_path(doc_name):
	"""Deterministic private path of an import's rollback manifest."""
	return frappe.get_site_path("private", "files", f"Rollback_{doc_name}.csv")


def read_manifest(doc_name):
	"""Returns [(action, doctype, name), ...] for everything this import wrote."""
	path = manifest_path(doc_name)
	if not os.path.exists(path):
		return []

	entries = []
	with open(path, encoding="utf-8", newline="") as f:
		for row in csv.reader(f):
			if len(row) >= 3 and row[0] in ("I", "U"):
				entries.append((row[0], row[1], row[2]))
	return entries


def rewrite_manifest(doc_name, entries):
	"""Replaces the manifest with `entries` (or deletes it when nothing is left)."""
	path = manifest_path(doc_name)
	if not entries:
		if os.path.exists(path):
			os.remove(path)
		return

	with open(path, "w", encoding="utf-8", newline="") as f:
		csv.writer(f).writerows(entries)


# -------------------------------------------------------------------------
# TEMPLATE GENERATION
# -------------------------------------------------------------------------


def build_template_workbook(target_doctype, include_optional=True):
	"""Builds an .xlsx template: sheet 1 has the headers to fill, sheet 2 documents
	every column (type, mandatory, link target, allowed options)."""
	meta = frappe.get_meta(target_doctype)
	mandatory, optional = [], []

	for field in meta.fields:
		if field.fieldtype in LAYOUT_FIELDTYPES or field.fieldtype == "Table":
			continue
		if field.read_only or getattr(field, "is_virtual", 0):
			continue
		(mandatory if field.reqd else optional).append(field)

	columns = mandatory + (optional if include_optional else [])

	wb = openpyxl.Workbook()
	ws = wb.active
	ws.title = target_doctype[:31]

	headers = [f.label or f.fieldname for f in columns]
	ws.append(headers)
	for idx, field in enumerate(columns, start=1):
		cell = ws.cell(row=1, column=idx)
		cell.font = Font(bold=True, color="FFFFFF")
		cell.fill = PatternFill("solid", fgColor="C0392B" if field.reqd else "34495E")
		ws.column_dimensions[cell.column_letter].width = max(14, min(38, len(str(cell.value)) + 6))
	ws.freeze_panes = "A2"

	guide = wb.create_sheet("Field Guide")
	guide.append(["Column Header", "Fieldname", "Type", "Mandatory", "Links To / Options"])
	for field in columns:
		options = ""
		if field.fieldtype == "Link":
			options = _("Existing {0} (name or title)").format(field.options)
		elif field.fieldtype == "Select" and field.options:
			options = ", ".join(field.options.split("\n")[:12])
		guide.append(
			[
				field.label or field.fieldname,
				field.fieldname,
				field.fieldtype,
				"Yes" if field.reqd else "No",
				options,
			]
		)
	for column_cells in guide.columns:
		guide.column_dimensions[column_cells[0].column_letter].width = 32
	for cell in guide[1]:
		cell.font = openpyxl.styles.Font(bold=True)

	stream = io.BytesIO()
	wb.save(stream)
	wb.close()
	return stream.getvalue(), len(mandatory), len(columns)


# -------------------------------------------------------------------------
# FRAPPE BACKGROUND JOB WRAPPER & API METHODS
# -------------------------------------------------------------------------


def _get_import_doc(doc_name, ptype="write"):
	doc = frappe.get_doc("Smart Data Import", doc_name)
	doc.check_permission(ptype)
	return doc


@frappe.whitelist()
def analyze_smart_import(doc_name):
	engine = SmartImportEngine(_get_import_doc(doc_name))
	engine.analyze_files_and_build_graph()
	return frappe.get_doc("Smart Data Import", doc_name)


@frappe.whitelist()
def start_smart_import(doc_name):
	doc = _get_import_doc(doc_name)

	if doc.status == "Processing":
		return {"status": "error", "message": _("This import is already running.")}
	if not doc.files:
		frappe.throw(_("Attach at least one Excel or CSV file before starting the import."))

	# One click is enough: analyze automatically if the graph is missing or stale.
	if doc.status in ("Pending", "Analyzing") or not doc.dependencies:
		SmartImportEngine(doc).analyze_files_and_build_graph()
		doc.reload()

	blocked = [row.idx for row in doc.files if row.status == "Failed" and not cint(row.total_rows)]
	if not doc.dependencies:
		frappe.throw(
			_("Nothing to import. Check the file rows for errors: {0}").format(
				", ".join(str(i) for i in blocked) or _("no readable rows found")
			)
		)

	frappe.enqueue(
		"smart_data_import.smart_import_engine.run_async_import",
		queue="long",
		timeout=7200,
		doc_name=doc_name,
	)
	frappe.db.set_value("Smart Data Import", doc_name, "status", "Processing", update_modified=False)
	frappe.db.commit()
	return {
		"status": "queued",
		"message": _("Import queued in the background for {0} rows.").format(doc.total_records),
	}


@frappe.whitelist()
def reset_import(doc_name):
	"""Clears counters and statuses so a finished or failed import can be re-run.

	The rollback manifest is deliberately kept: the records created by the previous
	run still exist, so it must stay possible to undo them.
	"""
	doc = _get_import_doc(doc_name)
	if doc.status in ("Processing", "Rolling Back"):
		frappe.throw(_("Cannot reset while the import is running."))

	doc.status = "Pending"
	doc.imported_records = 0
	doc.failed_records = 0
	doc.skipped_records = 0
	doc.rolled_back_records = 0
	doc.progress_percent = 0
	doc.execution_time_seconds = 0
	doc.error_log = ""
	doc.error_file = None
	doc.errors = []
	for row in doc.files:
		row.status = "Pending"
		row.error_log = ""
	for row in doc.dependencies:
		row.status = "Pending"
		row.processed_count = 0
	doc.save(ignore_permissions=True)
	frappe.db.commit()
	return {"status": "reset", "message": _("Import reset. Re-analyze and start again when ready.")}


@frappe.whitelist()
def get_import_preview(doc_name, sample_size=5):
	"""Returns, per attached file, exactly how columns will be mapped plus a few
	sample rows — so mistakes are visible before any record is written."""
	doc = _get_import_doc(doc_name, "read")
	engine = SmartImportEngine(doc)
	sample_size = min(20, max(1, cint(sample_size) or 5))
	result = []

	for row in doc.files:
		entry = {
			"idx": row.idx,
			"file": row.file,
			"doctype": row.doctype_name,
			"sheet": row.sheet_name,
			"total_rows": cint(row.total_rows),
			"error": None,
			"mapped": {},
			"unmapped": [],
			"missing_mandatory": [],
			"headers": [],
			"samples": [],
		}

		file_path, error = resolve_file_path(row.file)
		if error:
			entry["error"] = error
			result.append(entry)
			continue

		if not row.doctype_name:
			entry["error"] = _("Target DocType not set for this file.")
			result.append(entry)
			continue

		generator = engine._get_row_generator(file_path, row.sheet_name)
		headers = SmartImportEngine._clean_header_row(next(generator, []) or [])
		samples = []
		for values in generator:
			samples.append(["" if v is None else str(v) for v in values])
			if len(samples) >= sample_size:
				break

		meta = frappe.get_meta(row.doctype_name)
		column_map, defaults = engine.rules_for(row.doctype_name)
		mapping = map_columns(meta, headers, column_map)

		entry["headers"] = headers
		entry["samples"] = samples
		entry["mapped"] = mapping["mapped"]
		entry["unmapped"] = mapping["unmapped"]
		entry["missing_mandatory"] = missing_mandatory_fields(meta, mapping["mapped"].values(), defaults)
		entry["defaults"] = defaults
		result.append(entry)

	return {"files": result, "import_type": doc.import_type, "status": doc.status}


@frappe.whitelist()
def download_import_template(target_doctype, include_optional=1):
	"""Generates a ready-to-fill Excel template for a DocType and returns its URL."""
	if not frappe.db.exists("DocType", target_doctype):
		frappe.throw(_("DocType {0} does not exist.").format(target_doctype))
	frappe.has_permission(target_doctype, "create", throw=True)

	content, mandatory_count, total_columns = build_template_workbook(
		target_doctype, include_optional=cint(include_optional)
	)

	from frappe.utils.file_manager import save_file

	file_doc = save_file(
		f"Template_{target_doctype.replace(' ', '_')}.xlsx",
		content,
		None,
		None,
		is_private=1,
	)
	return {
		"file_url": file_doc.file_url,
		"file_name": file_doc.file_name,
		"mandatory_columns": mandatory_count,
		"total_columns": total_columns,
	}


def error_signature(reason):
	"""Collapses row-specific details so 496 failures group into a handful of causes.

	"Could not find Customer Group: BH-Vendor" and "...: Marketing -Suppliers" both
	become "Could not find Customer Group: …".
	"""
	text = str(reason or "").strip()
	text = re.sub(r'"[^"]*"', '"…"', text)
	text = re.sub(r"\b[A-Z][A-Z0-9]*-[\w-]*\d[\w-]*\b", "…", text)
	text = re.sub(r"(:\s*)(?!…).+$", r"\1…", text)
	text = re.sub(r"\s+", " ", text)
	return text[:200] or _("Unknown error")


@frappe.whitelist()
def get_error_summary(doc_name, limit=25):
	"""Groups the stored problem rows by cause, most frequent first — this is what
	the "View Errors" dialog shows."""
	doc = _get_import_doc(doc_name, "read")
	limit = min(100, max(1, cint(limit) or 25))

	error_rows = doc.get("errors") or []
	groups = {}
	for row in error_rows:
		key = (row.error_type or "Failed", error_signature(row.reason))
		group = groups.setdefault(
			key,
			{
				"type": key[0],
				"problem": key[1],
				"count": 0,
				"example_row": row.row_index,
				"example_reason": row.reason,
				"doctypes": set(),
			},
		)
		group["count"] += 1
		if row.doctype_name:
			group["doctypes"].add(row.doctype_name)

	ranked = sorted(groups.values(), key=lambda g: (-g["count"], g["problem"]))[:limit]
	for group in ranked:
		group["doctypes"] = ", ".join(sorted(group["doctypes"]))

	return {
		"groups": ranked,
		"shown_rows": len(error_rows),
		"failed": cint(doc.failed_records),
		"skipped": cint(doc.get("skipped_records")),
		"truncated": len(error_rows) >= ERROR_ROWS_IN_UI,
		"log_file": doc.error_file,
		"status": doc.status,
	}


@frappe.whitelist()
def get_rollback_summary(doc_name):
	"""What an "Undo Import" would delete, per DocType — read-only."""
	doc = _get_import_doc(doc_name, "read")
	entries = read_manifest(doc_name)

	per_doctype = {}
	for action, target_doctype, _name in entries:
		bucket = per_doctype.setdefault(target_doctype, {"inserted": 0, "updated": 0})
		bucket["inserted" if action == "I" else "updated"] += 1

	return {
		"deletable": sum(1 for e in entries if e[0] == "I"),
		"updated": sum(1 for e in entries if e[0] == "U"),
		"doctypes": [
			{"doctype": dt, "inserted": v["inserted"], "updated": v["updated"]}
			for dt, v in sorted(per_doctype.items())
		],
		"status": doc.status,
		"can_rollback": bool(entries) and doc.status not in ("Processing", "Rolling Back"),
	}


@frappe.whitelist()
def rollback_import(doc_name, force=0):
	"""Queues deletion of every record this import created.

	`force` skips Frappe's link validations — it can break documents that reference
	the deleted records, so it stays opt-in.
	"""
	doc = _get_import_doc(doc_name, "delete")
	if doc.status in ("Processing", "Rolling Back"):
		frappe.throw(_("Wait for the current job to finish before rolling back."))

	entries = read_manifest(doc_name)
	if not entries:
		frappe.throw(
			_("Nothing to roll back — no record of created documents exists for this import.")
		)
	if not any(e[0] == "I" for e in entries):
		frappe.throw(
			_("This import only updated existing records, which cannot be undone automatically.")
		)

	for target_doctype in {e[1] for e in entries if e[0] == "I"}:
		frappe.has_permission(target_doctype, "delete", throw=True)

	frappe.enqueue(
		"smart_data_import.smart_import_engine.run_async_rollback",
		queue="long",
		timeout=7200,
		doc_name=doc_name,
		force=cint(force),
	)
	frappe.db.set_value("Smart Data Import", doc_name, "status", "Rolling Back", update_modified=False)
	frappe.db.commit()
	return {
		"status": "queued",
		"message": _("Rollback queued: deleting {0} records in the background.").format(
			sum(1 for e in entries if e[0] == "I")
		),
	}


@frappe.whitelist()
def get_importable_doctypes(txt=""):
	"""Link-query friendly list of DocTypes that can be imported into."""
	filters = {"istable": 0, "issingle": 0}
	if txt:
		filters["name"] = ("like", f"%{txt}%")
	return frappe.get_all("DocType", filters=filters, pluck="name", order_by="name", limit_page_length=50)


def run_async_import(doc_name):
	"""Background worker entry point. Any crash must leave the document in a
	recoverable state, otherwise the UI stays stuck on 'Processing' forever."""
	engine = None
	try:
		engine = SmartImportEngine(doc_name)
		engine.execute_import()
	except Exception:
		traceback_text = frappe.get_traceback()
		frappe.db.rollback()
		frappe.log_error(title=f"Smart Data Import failed: {doc_name}", message=traceback_text)
		frappe.db.set_value(
			"Smart Data Import",
			doc_name,
			{
				"status": "Failed",
				"error_log": traceback_text[-5000:],
				"progress_percent": 100,
			},
			update_modified=False,
		)
		frappe.db.commit()
		if engine:
			engine.doc.status = "Failed"
			engine.publish_progress(_("Import failed. See the Detailed Error Log."), 100)
		raise


def run_async_rollback(doc_name, force=0):
	"""Background worker entry point for rollback. Like the import, a crash must not
	leave the document stuck on 'Rolling Back'."""
	engine = None
	try:
		engine = SmartImportEngine(doc_name)
		return engine.execute_rollback(force=bool(cint(force)))
	except Exception:
		traceback_text = frappe.get_traceback()
		frappe.db.rollback()
		frappe.log_error(title=f"Smart Data Import rollback failed: {doc_name}", message=traceback_text)
		frappe.db.set_value(
			"Smart Data Import",
			doc_name,
			{
				"status": "Failed",
				"error_log": traceback_text[-5000:],
				"progress_percent": 100,
			},
			update_modified=False,
		)
		frappe.db.commit()
		if engine:
			engine.doc.status = "Failed"
			engine.publish_progress(_("Rollback failed. See the Detailed Error Log."), 100)
		raise

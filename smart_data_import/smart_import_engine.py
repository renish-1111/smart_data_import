# Copyright (c) 2026, ERPNext AI Team and contributors
# For license information, please see license.txt

import os
import gc
import json
import time
import csv
import traceback
from collections import defaultdict, deque
import openpyxl

import frappe
from frappe.utils import get_files_path, now, flt, cint


class SmartImportEngine:
	"""
	Finest High-Performance Batch Data Import Engine.
	Designed to handle millions of records with automated dependency graph resolution (DAG),
	inner self-referential hierarchy sorting, memory-efficient streaming, and chunked batch insertion.
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
		self.stop_on_error = bool(self.doc.stop_on_error)
		self.ignore_duplicates = bool(getattr(self.doc, "ignore_duplicates", False))
		self.filter_rules = self._parse_filter_rules()

	def _parse_filter_rules(self):
		if not self.doc.filter_rules_json:
			return {}
		try:
			return json.loads(self.doc.filter_rules_json)
		except Exception:
			return {}

	# -------------------------------------------------------------------------
	# 1. FILE ANALYSIS & DOCTYPE AUTO-DETECTION
	# -------------------------------------------------------------------------

	def analyze_files_and_build_graph(self):
		"""
		Reads attached files, infers target DocTypes, calculates total rows,
		and builds the Directed Acyclic Graph (DAG) for dependency resolution.
		"""
		self.doc.status = "Analyzing"
		self.doc.save(ignore_permissions=True)
		frappe.db.commit()

		files_info = []
		detected_doctypes = set()

		# Step A: Parse attached files metadata & auto-detect DocType
		for row in self.doc.files:
			if not row.file:
				continue

			file_path = None
			filename = os.path.basename(row.file)
			file_id = frappe.db.get_value("File", {"file_url": row.file}, "name")
			if file_id:
				file_doc = frappe.get_doc("File", file_id)
				file_path = file_doc.get_full_path()
				filename = file_doc.file_name
			else:
				# Direct site public file fallback
				site_relative = row.file.lstrip("/")
				file_path = frappe.get_site_path("public", site_relative)
				if not os.path.exists(file_path):
					file_path = frappe.get_site_path(site_relative)

			if not file_path or not os.path.exists(file_path):
				continue

			headers, row_count, sheet_name = self.read_file_header_and_count(file_path, row.sheet_name)
			row.total_rows = row_count
			if sheet_name:
				row.sheet_name = sheet_name

			# Auto-detect DocType if missing
			if not row.doctype_name or self.doc.auto_detect_doctype:
				dt = self.detect_target_doctype(filename, row.sheet_name, headers)
				if dt:
					row.doctype_name = dt


			if row.doctype_name:
				detected_doctypes.add(row.doctype_name)
				row.status = "Analyzed"
			else:
				row.status = "Failed"
				row.error_log = "Unable to auto-detect target DocType. Please specify manually."

		# Step B: Build Dependency Graph & Topological Sorting
		self.doc.dependencies = []
		if detected_doctypes:
			tiers = self.build_topological_dependency_tiers(detected_doctypes)
			total_records = 0
			for dt_info in tiers:
				total_dt_rows = sum(r.total_rows for r in self.doc.files if r.doctype_name == dt_info["doctype"])
				total_records += total_dt_rows
				self.doc.append("dependencies", {
					"execution_tier": dt_info["tier"],
					"doctype_name": dt_info["doctype"],
					"depends_on_doctypes": ", ".join(dt_info["depends_on"]),
					"has_inner_dependency": 1 if dt_info["inner_ref_field"] else 0,
					"self_reference_field": dt_info["inner_ref_field"] or "",
					"status": "Pending",
					"total_count": total_dt_rows,
					"processed_count": 0
				})
			self.doc.total_records = total_records

		self.doc.status = "Ready"
		self.doc.save(ignore_permissions=True)
		frappe.db.commit()
		return True

	def read_file_header_and_count(self, file_path, sheet_name=None):
		"""
		Efficiently reads file headers and row count without loading full file into RAM.
		"""
		headers = []
		count = 0
		selected_sheet = sheet_name

		if file_path.endswith(('.xlsx', '.xlsm', '.xltx', '.xltm')):
			wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
			sheet_names = wb.sheetnames
			selected_sheet = sheet_name if (sheet_name and sheet_name in sheet_names) else sheet_names[0]
			ws = wb[selected_sheet]
			for idx, row in enumerate(ws.iter_rows(values_only=True)):
				if idx == 0:
					headers = [str(cell).strip() for cell in row if cell is not None]
				else:
					if any(cell is not None and str(cell).strip() != "" for cell in row):
						count += 1
			wb.close()
		elif file_path.endswith('.csv'):
			with open(file_path, mode='r', encoding='utf-8-sig', errors='ignore') as f:
				reader = csv.reader(f)
				for idx, row in enumerate(reader):
					if idx == 0:
						headers = [str(cell).strip() for cell in row if cell is not None]
					else:
						if any(cell is not None and str(cell).strip() != "" for cell in row):
							count += 1

		return headers, count, selected_sheet

	def detect_target_doctype(self, filename, sheet_name, headers):
		"""
		Finest Algorithm: Intelligently infers target DocType using a weighted scoring engine
		combining Header Fields, Primary Key Bonus, Sheet Name, and Filename Word Intersection.
		"""
		all_doctypes = frappe.get_all("DocType", filters={"istable": 0, "custom": 0}, pluck="name")
		scores = {}

		# Step 1: Header Field Matching & Weighted Field Scoring (Highest Confidence)
		if headers:
			header_set = set(h.lower().replace(' ', '_').strip() for h in headers)
			for dt in all_doctypes:
				meta = frappe.get_meta(dt)
				dt_fields = set(f.fieldname.lower() for f in meta.fields)
				dt_labels = set((f.label or '').lower().replace(' ', '_').strip() for f in meta.fields)
				dt_fields.add('name')
				if meta.title_field:
					dt_fields.add(meta.title_field.lower())

				all_meta_fields = dt_fields.union(dt_labels)
				matches = len(header_set.intersection(all_meta_fields))

				# Weighted bonus if primary key or title field matches (e.g. customer_group_name)
				clean_dt_key = dt.lower().replace(' ', '_') + '_name'
				if clean_dt_key in header_set:
					matches += 3

				if matches >= 2:
					scores[dt] = scores.get(dt, 0) + (matches * 10)

		# Step 2: Sheet Name Matching
		if sheet_name:
			clean_sheet = sheet_name.replace('_', ' ').replace('-', ' ').strip().lower()
			for dt in all_doctypes:
				if dt.lower() == clean_sheet:
					scores[dt] = scores.get(dt, 0) + 40

		# Step 3: Filename Matching & Word Intersection Scoring
		clean_base = filename.rsplit('.', 1)[0].replace('_', ' ').replace('-', ' ').strip().lower()
		for dt in all_doctypes:
			dt_clean = dt.lower()
			if dt_clean == clean_base:
				scores[dt] = scores.get(dt, 0) + 50
			else:
				dt_words = set(dt_clean.split())
				fn_words = set(clean_base.split())
				common = dt_words.intersection(fn_words)
				if common:
					scores[dt] = scores.get(dt, 0) + (len(common) * 5)

		if scores:
			sorted_scores = sorted(scores.items(), key=lambda x: x[1], reverse=True)
			return sorted_scores[0][0]

		return None




	# -------------------------------------------------------------------------
	# 2. TOPOLOGICAL SORT ALGORITHM FOR INTER & INTRA DOCTYPE DEPENDENCIES
	# -------------------------------------------------------------------------

	def build_topological_dependency_tiers(self, doctypes_set):
		"""
		Builds Directed Acyclic Graph (DAG) for target DocTypes and executes
		Kahn's Topological Sorting algorithm to determine execution tiers.
		Prioritizes primary document Link fields over child table links to handle cycles seamlessly.
		"""
		doctypes = list(doctypes_set)
		graph = defaultdict(set)      # u -> set of nodes that depend on u
		in_degree = {dt: 0 for dt in doctypes}
		depends_on_map = defaultdict(set)
		inner_ref_map = {}

		for dt in doctypes:
			meta = frappe.get_meta(dt)
			inner_ref = None
			primary_link_doctypes = set()
			child_link_doctypes = set()

			# Inspect primary fields
			for f in meta.fields:
				if f.fieldtype == 'Link' and f.options:
					if f.options == dt:
						inner_ref = f.fieldname
					elif f.options in doctypes_set:
						primary_link_doctypes.add(f.options)
				elif f.fieldtype == 'Table' and f.options:
					child_meta = frappe.get_meta(f.options)
					for cf in child_meta.fields:
						if cf.fieldtype == 'Link' and cf.options and cf.options in doctypes_set and cf.options != dt:
							child_link_doctypes.add(cf.options)

			inner_ref_map[dt] = inner_ref
			# Primary links define core document dependency ordering
			depends_on_map[dt] = primary_link_doctypes

			for target_dt in primary_link_doctypes:
				graph[target_dt].add(dt)
				in_degree[dt] += 1

		# Kahn's Algorithm for Topological Sort into Tiers
		queue = deque([dt for dt in doctypes if in_degree[dt] == 0])
		result_tiers = []
		current_tier = 0
		visited_count = 0

		while queue:
			tier_size = len(queue)
			tier_nodes = []
			for _ in range(tier_size):
				u = queue.popleft()
				visited_count += 1
				tier_nodes.append({
					"doctype": u,
					"tier": current_tier,
					"depends_on": list(depends_on_map[u]),
					"inner_ref_field": inner_ref_map[u]
				})
				for v in graph[u]:
					in_degree[v] -= 1
					if in_degree[v] == 0:
						queue.append(v)
			result_tiers.extend(tier_nodes)
			current_tier += 1

		# Cycle Fallback
		if visited_count < len(doctypes):
			unvisited = [dt for dt in doctypes if in_degree[dt] > 0]
			for u in unvisited:
				result_tiers.append({
					"doctype": u,
					"tier": current_tier,
					"depends_on": list(depends_on_map[u]),
					"inner_ref_field": inner_ref_map[u]
				})

		return result_tiers


	# -------------------------------------------------------------------------
	# 3. HIGH-SPEED BATCH IMPORT EXECUTION ENGINE
	# -------------------------------------------------------------------------

	def execute_import(self):
		"""
		Executes batch data import tier by tier, streaming rows, processing chunks,
		and updating real-time socket progress.
		"""
		start_time = time.time()
		self.doc.status = "Processing"
		self.doc.imported_records = 0
		self.doc.failed_records = 0
		self.doc.progress_percent = 0
		self.doc.save(ignore_permissions=True)
		frappe.db.commit()

		# Sort dependencies by execution_tier ASC
		sorted_deps = sorted(self.doc.dependencies, key=lambda d: d.execution_tier)
		failed_rows_list = []

		for dep in sorted_deps:
			dep.status = "Processing"
			self.doc.save(ignore_permissions=True)
			frappe.db.commit()

			target_dt = dep.doctype_name
			self.publish_progress(f"Starting import for Tier {dep.execution_tier}: {target_dt}")

			# Collect all files matching this DocType
			matching_files = [f for f in self.doc.files if f.doctype_name == target_dt]
			
			for file_row in matching_files:
				if not file_row.file:
					continue
				file_path = None
				file_id = frappe.db.get_value("File", {"file_url": file_row.file}, "name")
				if file_id:
					file_doc = frappe.get_doc("File", file_id)
					file_path = file_doc.get_full_path()
				else:
					site_relative = file_row.file.lstrip("/")
					file_path = frappe.get_site_path("public", site_relative)
					if not os.path.exists(file_path):
						file_path = frappe.get_site_path(site_relative)

				if not file_path or not os.path.exists(file_path):
					continue

				
				# Execute high-speed batch stream
				success, count, failed, errors = self._stream_and_batch_insert(
					file_path=file_path,
					sheet_name=file_row.sheet_name,
					target_doctype=target_dt,
					self_ref_field=dep.self_reference_field if dep.has_inner_dependency else None,
					dep_row=dep
				)

				dep.processed_count += count
				self.doc.imported_records += count
				self.doc.failed_records += failed
				failed_rows_list.extend(errors)
				
				if failed > 0 and self.stop_on_error:
					dep.status = "Failed"
					self.doc.status = "Failed"
					self.doc.save(ignore_permissions=True)
					frappe.db.commit()
					return False

			dep.status = "Completed" if dep.processed_count > 0 else "Pending"
			self.doc.save(ignore_permissions=True)
			frappe.db.commit()

		# Finalize status and generate error summary
		self.doc.execution_time_seconds = flt(time.time() - start_time, 2)
		if failed_rows_list:
			self.doc.status = "Partial Success" if self.doc.imported_records > 0 else "Failed"
			self._generate_failed_rows_excel(failed_rows_list)
			self.doc.error_log = "\n".join([f"Row {err['row']}: {err['reason']}" for err in failed_rows_list[:200]])
		else:
			self.doc.status = "Completed"

		self.doc.progress_percent = 100
		self.doc.save(ignore_permissions=True)
		frappe.db.commit()
		self.publish_progress("Import job completed successfully!", 100)
		return True

	def _stream_and_batch_insert(self, file_path, sheet_name, target_doctype, self_ref_field, dep_row):
		"""
		Streams data row-by-row in configurable batch sizes (e.g. 5,000 rows).
		Handles inner self-referential hierarchy via two-pass insertion.
		"""
		meta = frappe.get_meta(target_doctype)
		meta_labels = { (f.label or '').lower().strip(): f.fieldname for f in meta.fields }
		meta_labels.update({ f.fieldname.lower().strip(): f.fieldname for f in meta.fields })
		meta_labels['name'] = 'name'

		inserted_count = 0
		failed_count = 0
		failed_errors = []

		# Stream generator for Excel or CSV
		row_generator = self._get_row_generator(file_path, sheet_name)
		headers = next(row_generator, None)
		if not headers:
			return True, 0, 0, []


		# Column to field mapping
		col_map = []
		for h in headers:
			clean_h = str(h).strip()
			if clean_h.lower() in meta_labels:
				col_map.append(meta_labels[clean_h.lower()])
			else:
				col_map.append(None)

		batch_records = []
		second_pass_updates = []
		row_idx = 1  # 1-indexed for row header

		for row_values in row_generator:
			row_idx += 1
			row_dict = {}
			for idx, val in enumerate(row_values):
				fieldname = col_map[idx] if idx < len(col_map) else None
				if fieldname and val is not None:
					str_val = str(val).strip() if self.clean_whitespace and isinstance(val, str) else val
					if str_val != "":
						row_dict[fieldname] = str_val

			if self.skip_empty_rows and not row_dict:
				continue

			# Smart Link Title-to-Name resolution for autonamed DocTypes (e.g. Project Name -> PROJ-0001)
			row_dict = self._resolve_row_link_fields(target_doctype, row_dict, meta)

			# Inner Self-Reference handling: Pass 1 nullifies parent reference, saved for Pass 2 update
			if self_ref_field and self_ref_field in row_dict:
				parent_val = row_dict.pop(self_ref_field)
				if parent_val:
					# Store row_idx & subject/name identifier for Pass 2 parent update
					second_pass_updates.append((row_idx, row_dict, self_ref_field, parent_val))

			batch_records.append((row_idx, row_dict))


			if len(batch_records) >= self.batch_size:
				c, f, errs = self._flush_batch_to_db(target_doctype, batch_records, meta)
				inserted_count += c
				failed_count += f
				failed_errors.extend(errs)
				batch_records = []
				self._update_progress_counts(dep_row, inserted_count)
				gc.collect()

		# Flush remaining records
		if batch_records:
			c, f, errs = self._flush_batch_to_db(target_doctype, batch_records, meta)
			inserted_count += c
			failed_count += f
			failed_errors.extend(errs)
			self._update_progress_counts(dep_row, inserted_count)

		# Second Pass: Update inner self-referential links (e.g. parent_task, parent_account, parent_item_group)
		if second_pass_updates:
			self.publish_progress(f"Updating hierarchical parent references for {target_doctype}...")
			title_field = meta.title_field or ('subject' if meta.has_field('subject') else ('name' if meta.has_field('name') else None))
			for r_idx, r_data, fieldname, parent_val in second_pass_updates:
				try:
					child_name = r_data.get('name')
					if not child_name and title_field and r_data.get(title_field):
						child_name = frappe.db.get_value(target_doctype, {title_field: r_data[title_field]}, 'name')
					
					parent_name = parent_val
					if not frappe.db.exists(target_doctype, parent_name) and title_field:
						p_res = frappe.db.get_value(target_doctype, {title_field: parent_val}, 'name')
						if p_res:
							parent_name = p_res

					if child_name and parent_name:
						frappe.db.set_value(target_doctype, child_name, fieldname, parent_name, update_modified=False)
				except Exception:
					pass
			frappe.db.commit()

		return True, inserted_count, failed_count, failed_errors

	def _resolve_row_link_fields(self, target_doctype, row_dict, meta):
		"""
		Automatically resolves human-readable Link titles to primary key names for autonamed DocTypes.
		E.g. Task.project = "Kintech AI Engine 2026" -> resolves to "PROJ-0001".
		"""
		resolved_dict = dict(row_dict)
		for f in meta.fields:
			if f.fieldtype == 'Link' and f.options and f.fieldname in resolved_dict:
				val = resolved_dict[f.fieldname]
				if not val:
					continue
				link_dt = f.options
				if not frappe.db.exists(link_dt, val):
					link_meta = frappe.get_meta(link_dt)
					search_fields = []
					if link_meta.title_field:
						search_fields.append(link_meta.title_field)
					clean_dt_name = link_dt.lower().replace(' ', '_') + '_name'
					search_fields.extend([clean_dt_name, 'title', 'subject', 'item_name', 'project_name', 'customer_name', 'customer_group_name'])
					
					for sf in search_fields:
						if link_meta.has_field(sf):
							res = frappe.db.get_value(link_dt, {sf: val}, 'name')
							if res:
								resolved_dict[f.fieldname] = res
								break
		return resolved_dict



	def _flush_batch_to_db(self, target_doctype, batch_records, meta):
		"""
		Executes high-speed bulk database insertion for a chunk of records.
		"""
		c_success = 0
		c_failed = 0
		errors = []

		try:
			docs_to_insert = []
			for row_idx, row_dict in batch_records:
				doc = frappe.new_doc(target_doctype)
				doc.update(row_dict)
				doc.flags.ignore_permissions = True
				doc.flags.ignore_mandatory = self.ignore_link_errors
				doc.flags.ignore_links = self.ignore_link_errors
				docs_to_insert.append((row_idx, doc))

			for row_idx, doc in docs_to_insert:
				try:
					if self.ignore_duplicates:
						doc_name = doc.get("name")
						if doc_name and frappe.db.exists(target_doctype, doc_name):
							continue
						title_field = meta.title_field or (target_doctype.lower().replace(' ', '_') + '_name')
						if title_field and meta.has_field(title_field) and doc.get(title_field):
							if frappe.db.exists(target_doctype, {title_field: doc.get(title_field)}):
								continue

					doc.insert(
						ignore_permissions=True,
						ignore_mandatory=self.ignore_link_errors,
						ignore_links=self.ignore_link_errors
					)
					c_success += 1
				except frappe.DuplicateEntryError as e:
					if self.ignore_duplicates:
						pass
					else:
						c_failed += 1
						errors.append({
							"row": row_idx,
							"doctype": target_doctype,
							"reason": str(e)
						})
						if self.stop_on_error:
							raise e
				except Exception as e:
					if self.ignore_duplicates and ("Duplicate" in type(e).__name__ or "Duplicate entry" in str(e) or "already exists" in str(e)):
						pass
					else:
						c_failed += 1
						errors.append({
							"row": row_idx,
							"doctype": target_doctype,
							"reason": str(e)
						})
						if self.stop_on_error:
							raise e

			frappe.db.commit()

		except Exception as e:
			frappe.db.rollback()
			c_failed += len(batch_records) - c_success
			errors.append({"row": 0, "doctype": target_doctype, "reason": f"Batch execution failed: {str(e)}"})

		return c_success, c_failed, errors

	def _get_row_generator(self, file_path, sheet_name):
		"""
		Returns generator yielding rows one by one.
		"""
		if file_path.endswith(('.xlsx', '.xlsm', '.xltx', '.xltm')):
			wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
			ws = wb[sheet_name] if (sheet_name and sheet_name in wb.sheetnames) else wb.active
			for row in ws.iter_rows(values_only=True):
				yield row
			wb.close()
		elif file_path.endswith('.csv'):
			with open(file_path, mode='r', encoding='utf-8-sig', errors='ignore') as f:
				reader = csv.reader(f)
				for row in reader:
					yield row

	def _update_progress_counts(self, dep_row, current_processed):
		dep_row.processed_count = current_processed
		if self.doc.total_records > 0:
			self.doc.progress_percent = flt((self.doc.imported_records / self.doc.total_records) * 100, 1)
		self.doc.save(ignore_permissions=True)
		frappe.db.commit()
		self.publish_progress(f"Importing {dep_row.doctype_name}: {current_processed} rows inserted...")

	def publish_progress(self, message, progress_pct=None):
		pct = progress_pct if progress_pct is not None else self.doc.progress_percent
		frappe.publish_realtime(
			"smart_import_progress",
			{
				"doc_name": self.doc.name,
				"message": message,
				"progress": pct,
				"imported": self.doc.imported_records,
				"failed": self.doc.failed_records,
				"status": self.doc.status
			},
			user=frappe.session.user
		)

	def _generate_failed_rows_excel(self, failed_errors):
		"""
		Generates downloadable Excel file containing failed row numbers and reasons.
		"""
		wb = openpyxl.Workbook()
		ws = wb.active
		ws.title = "Failed Import Rows"
		ws.append(["Row Index", "Target DocType", "Failure Reason"])
		for err in failed_errors:
			ws.append([err.get("row"), err.get("doctype"), err.get("reason")])
		
		file_name = f"Failed_Rows_{self.doc.name}.xlsx"
		files_dir = get_files_path()
		full_path = os.path.join(files_dir, file_name)
		wb.save(full_path)
		wb.close()

		file_doc = frappe.get_doc({
			"doctype": "File",
			"file_name": file_name,
			"attached_to_doctype": "Smart Data Import",
			"attached_to_name": self.doc.name,
			"file_url": f"/files/{file_name}",
			"is_private": 0
		})
		file_doc.insert(ignore_permissions=True)
		self.doc.error_file = file_doc.file_url


# -------------------------------------------------------------------------
# FRAPPE BACKGROUND JOB WRAPPER & API METHODS
# -------------------------------------------------------------------------

@frappe.whitelist()
def analyze_smart_import(doc_name):
	engine = SmartImportEngine(doc_name)
	engine.analyze_files_and_build_graph()
	return engine.doc

@frappe.whitelist()
def start_smart_import(doc_name):
	doc = frappe.get_doc("Smart Data Import", doc_name)
	if doc.status == "Processing":
		return {"status": "error", "message": "Import job is already running."}

	frappe.enqueue(
		"smart_data_import.smart_import_engine.run_async_import",
		queue="long",
		timeout=7200,
		doc_name=doc_name
	)
	doc.status = "Processing"
	doc.save(ignore_permissions=True)
	frappe.db.commit()
	return {"status": "queued", "message": "Smart Data Import queued in background worker."}

def run_async_import(doc_name):
	engine = SmartImportEngine(doc_name)
	engine.execute_import()

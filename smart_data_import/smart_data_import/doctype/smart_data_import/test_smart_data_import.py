# Copyright (c) 2026, ERPNext AI Team and contributors
# For license information, please see license.txt

import frappe
from frappe.tests import IntegrationTestCase
from frappe.utils.file_manager import save_file

from smart_data_import.smart_import_engine import (
	SmartImportEngine,
	build_template_workbook,
	error_signature,
	get_error_summary,
	get_import_preview,
	get_rollback_summary,
	map_columns,
	missing_mandatory_fields,
	normalize_key,
	read_manifest,
	reset_import,
	rollback_import,
)

NOTE_CSV = """Title,Weird Body Column,Public
{a},body one,1
{b},body two,0
"""


class TestSmartDataImport(IntegrationTestCase):
	def setUp(self):
		frappe.set_user("Administrator")
		self.suffix = frappe.generate_hash(length=6)
		self.created_notes = []
		self.import_doc = None

	def tearDown(self):
		for name in self.created_notes:
			frappe.delete_doc("Note", name, force=True, ignore_permissions=True, delete_permanently=True)
		if self.import_doc:
			frappe.delete_doc("Smart Data Import", self.import_doc, force=True, ignore_permissions=True)
		frappe.db.commit()

	# ------------------------------------------------------------------ helpers

	def make_import(self, filename, content, **kwargs):
		doc = frappe.get_doc(dict({"doctype": "Smart Data Import", "title": f"Test SDI {self.suffix}"}, **kwargs))
		doc.insert(ignore_permissions=True)
		self.import_doc = doc.name

		file_doc = save_file(filename, content.encode(), "Smart Data Import", doc.name, is_private=0)
		doc.append("files", {"file": file_doc.file_url})
		doc.save(ignore_permissions=True)
		frappe.db.commit()
		return doc

	def note_csv(self):
		return NOTE_CSV.format(a=f"SDI Test A {self.suffix}", b=f"SDI Test B {self.suffix}")

	def collect_notes(self):
		self.created_notes = frappe.get_all(
			"Note", filters={"title": ("like", f"SDI Test%{self.suffix}")}, pluck="name"
		)
		return self.created_notes

	# ------------------------------------------------------------------ mapping

	def test_normalize_key(self):
		self.assertEqual(normalize_key("Customer  Name"), "customer name")
		self.assertEqual(normalize_key("customer_name"), "customer name")
		self.assertEqual(normalize_key("Customer-Name"), "customer name")

	def test_map_columns_labels_fieldnames_and_overrides(self):
		meta = frappe.get_meta("Note")
		mapping = map_columns(meta, ["Title", "content", "Nonsense"], {"Nonsense": "public"})
		self.assertEqual(mapping["mapped"]["Title"], "title")
		self.assertEqual(mapping["mapped"]["content"], "content")
		self.assertEqual(mapping["mapped"]["Nonsense"], "public")
		self.assertEqual(mapping["unmapped"], [])

	def test_map_columns_keeps_positions_for_blank_headers(self):
		meta = frappe.get_meta("Note")
		mapping = map_columns(meta, ["Title", "", "content"], None)
		self.assertEqual(mapping["col_map"], ["title", None, "content"])

	def test_missing_mandatory_fields(self):
		meta = frappe.get_meta("Note")
		self.assertIn("Title", missing_mandatory_fields(meta, ["content"]))
		self.assertNotIn("Title", missing_mandatory_fields(meta, ["title", "content"]))
		self.assertNotIn("Title", missing_mandatory_fields(meta, ["content"], {"title": "x"}))

	# ------------------------------------------------------------------ detection & DAG

	def test_detect_target_doctype(self):
		engine = SmartImportEngine.__new__(SmartImportEngine)
		engine._dt_index = None
		detected, _note = engine.detect_target_doctype("Project.csv", None, ["Project Name", "Status"])
		self.assertEqual(detected, "Project")

	def test_dependency_tiers_and_self_reference(self):
		engine = SmartImportEngine.__new__(SmartImportEngine)
		tiers = engine.build_topological_dependency_tiers({"Project", "Task"})
		tier_map = {t["doctype"]: t["tier"] for t in tiers}
		self.assertLess(tier_map["Project"], tier_map["Task"])

		task_info = next(t for t in tiers if t["doctype"] == "Task")
		self.assertEqual(task_info["inner_ref_field"], "parent_task")

	# ------------------------------------------------------------------ validation

	def test_unsupported_file_is_rejected(self):
		doc = frappe.get_doc(
			{
				"doctype": "Smart Data Import",
				"title": f"Test SDI reject {self.suffix}",
				"files": [{"file": "/files/some_report.pdf"}],
			}
		)
		self.assertRaises(frappe.ValidationError, doc.insert)

	def test_invalid_mapping_json_is_rejected(self):
		doc = frappe.get_doc(
			{
				"doctype": "Smart Data Import",
				"title": f"Test SDI json {self.suffix}",
				"filter_rules_json": "{not json",
			}
		)
		self.assertRaises(frappe.ValidationError, doc.insert)

	# ------------------------------------------------------------------ end to end

	def test_end_to_end_import_with_override_and_duplicate_skip(self):
		doc = self.make_import(
			"Note.csv",
			self.note_csv(),
			filter_rules_json='{"Note": {"column_map": {"Weird Body Column": "content"}}}',
			ignore_duplicates=1,
		)

		SmartImportEngine(doc.name).analyze_files_and_build_graph()
		doc.reload()
		self.assertEqual(doc.files[0].doctype_name, "Note")
		self.assertEqual(doc.files[0].total_rows, 2)
		self.assertEqual(doc.status, "Ready")

		preview = get_import_preview(doc.name, 2)
		self.assertEqual(preview["files"][0]["mapped"]["Weird Body Column"], "content")
		self.assertEqual(len(preview["files"][0]["samples"]), 2)

		SmartImportEngine(doc.name).execute_import()
		doc.reload()
		self.assertEqual(doc.status, "Completed")
		self.assertEqual(doc.imported_records, 2)
		self.assertEqual(doc.failed_records, 0)

		notes = self.collect_notes()
		self.assertEqual(len(notes), 2)
		self.assertTrue(all(frappe.db.get_value("Note", n, "content") for n in notes))

		# Re-running must skip instead of failing or duplicating.
		reset_import(doc.name)
		SmartImportEngine(doc.name).execute_import()
		doc.reload()
		self.assertEqual(doc.imported_records, 0)
		self.assertEqual(doc.skipped_records, 2)
		self.assertEqual(doc.failed_records, 0)
		self.assertEqual(len(self.collect_notes()), 2)

	def test_missing_file_is_reported_on_the_row(self):
		doc = frappe.get_doc(
			{
				"doctype": "Smart Data Import",
				"title": f"Test SDI missing {self.suffix}",
				"files": [{"file": "/files/sdi_missing_file.csv", "doctype_name": "Note"}],
			}
		).insert(ignore_permissions=True)
		self.import_doc = doc.name

		SmartImportEngine(doc.name).analyze_files_and_build_graph()
		doc.reload()
		self.assertEqual(doc.files[0].status, "Failed")
		self.assertIn("could not be found", doc.files[0].error_log.lower())
		self.assertEqual(doc.status, "Pending")

	# ------------------------------------------------------------------ errors in the UI

	def test_error_signature_groups_row_specific_details(self):
		self.assertEqual(
			error_signature("Could not find Customer Group: BH-Vendor"),
			error_signature("Could not find Customer Group: Marketing -Suppliers"),
		)
		self.assertNotEqual(
			error_signature("Could not find Customer Group: BH-Vendor"),
			error_signature("Could not find Supplier Group: BH-Vendor"),
		)

	def test_failed_rows_are_listed_on_the_document(self):
		# 'Public' is not a valid ToDo field and priority is invalid -> every row fails.
		csv_content = (
			"Description,Priority\n"
			f"SDI Err A {self.suffix},Nonsense Priority\n"
			f"SDI Err B {self.suffix},Another Bad Priority\n"
		)
		doc = self.make_import("ToDo.csv", csv_content)
		doc.files[0].doctype_name = "ToDo"
		doc.save(ignore_permissions=True)

		SmartImportEngine(doc.name).analyze_files_and_build_graph()
		SmartImportEngine(doc.name).execute_import()
		doc.reload()

		self.assertEqual(doc.failed_records, 2)
		self.assertEqual(doc.status, "Failed")

		# The reasons must be readable on the form, not only in the attachment.
		self.assertEqual(len(doc.errors), 2)
		self.assertEqual(doc.errors[0].error_type, "Failed")
		self.assertEqual(doc.errors[0].doctype_name, "ToDo")
		self.assertEqual(doc.errors[0].row_index, 2)
		self.assertTrue(doc.errors[0].reason)
		self.assertNotIn("<", doc.errors[0].reason)  # HTML stripped for display

		summary = get_error_summary(doc.name)
		self.assertEqual(summary["failed"], 2)
		self.assertEqual(sum(g["count"] for g in summary["groups"]), 2)
		self.assertEqual(len(summary["groups"]), 1)  # same cause -> one group
		self.assertFalse(summary["truncated"])
		self.assertTrue(summary["log_file"])

		# Resetting clears them again.
		reset_import(doc.name)
		doc.reload()
		self.assertEqual(len(doc.errors), 0)

	def test_skipped_rows_are_listed_on_the_document(self):
		doc = self.make_import("Note.csv", self.note_csv(), ignore_duplicates=1)
		SmartImportEngine(doc.name).analyze_files_and_build_graph()
		SmartImportEngine(doc.name).execute_import()
		self.collect_notes()

		reset_import(doc.name)
		SmartImportEngine(doc.name).execute_import()
		doc.reload()

		self.assertEqual(doc.skipped_records, 2)
		self.assertEqual([row.error_type for row in doc.errors], ["Skipped", "Skipped"])
		summary = get_error_summary(doc.name)
		self.assertEqual(summary["skipped"], 2)
		self.assertEqual(summary["groups"][0]["type"], "Skipped")

	def test_rollback_deletes_created_records(self):
		doc = self.make_import(
			"Note.csv",
			self.note_csv(),
			filter_rules_json='{"Note": {"column_map": {"Weird Body Column": "content"}}}',
		)
		SmartImportEngine(doc.name).analyze_files_and_build_graph()
		SmartImportEngine(doc.name).execute_import()
		doc.reload()
		self.assertEqual(doc.imported_records, 2)
		self.assertEqual(len(self.collect_notes()), 2)

		# The manifest records what was created, so the run can be undone.
		manifest = read_manifest(doc.name)
		self.assertEqual(len([e for e in manifest if e[0] == "I"]), 2)
		self.assertTrue(doc.rollback_file)

		summary = get_rollback_summary(doc.name)
		self.assertEqual(summary["deletable"], 2)
		self.assertEqual(summary["doctypes"][0]["doctype"], "Note")

		result = SmartImportEngine(doc.name).execute_rollback()
		doc.reload()
		self.assertEqual(result["deleted"], 2)
		self.assertEqual(doc.status, "Rolled Back")
		self.assertEqual(doc.rolled_back_records, 2)
		self.assertEqual(self.collect_notes(), [])

		# Manifest is consumed, so there is nothing left to undo.
		self.assertEqual(read_manifest(doc.name), [])
		self.assertFalse(doc.rollback_file)
		self.assertRaises(frappe.ValidationError, rollback_import, doc.name)

	def test_rollback_refuses_when_only_updates_were_made(self):
		doc = self.make_import("Note.csv", self.note_csv(), import_type="Update Existing Records")
		SmartImportEngine(doc.name).analyze_files_and_build_graph()
		SmartImportEngine(doc.name).execute_import()
		doc.reload()

		# Nothing matched, so nothing was written and there is nothing to undo.
		self.assertEqual(doc.imported_records, 0)
		self.assertEqual(doc.skipped_records, 2)
		self.assertRaises(frappe.ValidationError, rollback_import, doc.name)

	def test_rollback_is_repeatable_after_partial_deletion(self):
		doc = self.make_import("Note.csv", self.note_csv())
		SmartImportEngine(doc.name).analyze_files_and_build_graph()
		SmartImportEngine(doc.name).execute_import()
		doc.reload()
		notes = self.collect_notes()
		self.assertEqual(len(notes), 2)

		# A record deleted by hand must not make the rollback fail.
		frappe.delete_doc("Note", notes[0], force=True, ignore_permissions=True, delete_permanently=True)
		frappe.db.commit()

		result = SmartImportEngine(doc.name).execute_rollback()
		self.assertEqual(result["deleted"], 2)  # the missing one counts as already undone
		self.assertEqual(self.collect_notes(), [])

	def test_template_contains_mandatory_columns_first(self):
		content, mandatory_count, total_columns = build_template_workbook("Note", include_optional=False)
		self.assertTrue(content.startswith(b"PK"))
		self.assertEqual(mandatory_count, total_columns)
		self.assertGreaterEqual(mandatory_count, 1)

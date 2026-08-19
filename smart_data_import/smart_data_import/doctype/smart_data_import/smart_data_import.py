# Copyright (c) 2026, ERPNext AI Team and contributors
# For license information, please see license.txt

import json
import os

import frappe
from frappe import _
from frappe.model.document import Document
from frappe.utils import cint

from smart_data_import.smart_import_engine import (
	SUPPORTED_EXTENSIONS,
	SmartImportEngine,
	download_import_template,
	get_error_summary,
	get_import_preview,
	get_rollback_summary,
	reset_import,
	rollback_import,
	start_smart_import,
)


class SmartDataImport(Document):
	def validate(self):
		if not self.title:
			self.title = f"Data Import - {frappe.utils.nowdate()}"

		if cint(self.batch_size) and cint(self.batch_size) < 500:
			self.batch_size = 500

		self.validate_files()
		self.validate_mapping_rules()

	def validate_files(self):
		"""Catches unusable attachments while the user is still on the form."""
		for row in self.files:
			if not row.file:
				continue
			extension = os.path.splitext(row.file.split("?")[0])[1].lower()
			if extension not in SUPPORTED_EXTENSIONS:
				frappe.throw(
					_("Row {0}: {1} files are not supported. Attach an Excel (.xlsx, .xlsm) or .csv file.").format(
						row.idx, extension or _("unknown")
					)
				)

	def validate_mapping_rules(self):
		"""Fails fast on malformed mapping JSON instead of silently ignoring it."""
		if not self.filter_rules_json or not self.filter_rules_json.strip():
			return
		try:
			rules = json.loads(self.filter_rules_json)
		except Exception as e:
			frappe.throw(_("Column Mapping & Defaults is not valid JSON: {0}").format(str(e)))

		if not isinstance(rules, dict):
			frappe.throw(_("Column Mapping & Defaults must be an object keyed by DocType name."))

		for doctype_name, config in rules.items():
			if not isinstance(config, dict):
				frappe.throw(_("Mapping rules for {0} must be an object.").format(doctype_name))
			for key in config:
				if key not in ("column_map", "defaults"):
					frappe.throw(
						_("Unknown mapping option {0} for {1}. Use 'column_map' or 'defaults'.").format(
							key, doctype_name
						)
					)

	@frappe.whitelist()
	def analyze_dependencies(self):
		engine = SmartImportEngine(self)
		engine.analyze_files_and_build_graph()
		# Reload from DB so client timestamps stay in sync.
		return frappe.get_doc("Smart Data Import", self.name)

	@frappe.whitelist()
	def start_import(self):
		return start_smart_import(self.name)

	@frappe.whitelist()
	def preview_mapping(self, sample_size=5):
		return get_import_preview(self.name, sample_size)

	@frappe.whitelist()
	def reset_for_rerun(self):
		return reset_import(self.name)

	@frappe.whitelist()
	def get_template(self, target_doctype, include_optional=1, mode="single"):
		return download_import_template(target_doctype, include_optional, mode)

	@frappe.whitelist()
	def error_summary(self, limit=25):
		return get_error_summary(self.name, limit)

	@frappe.whitelist()
	def rollback_summary(self):
		return get_rollback_summary(self.name)

	@frappe.whitelist()
	def rollback(self, force=0):
		return rollback_import(self.name, force)

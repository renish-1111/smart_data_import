# Copyright (c) 2026, ERPNext AI Team and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document
from smart_data_import.smart_import_engine import SmartImportEngine, start_smart_import


class SmartDataImport(Document):

	def validate(self):
		if not self.title:
			self.title = f"Data Import - {frappe.utils.nowdate()}"

	@frappe.whitelist()
	def analyze_dependencies(self):
		engine = SmartImportEngine(self)
		engine.analyze_files_and_build_graph()
		# Reload latest doc from DB to ensure timestamps are perfectly synchronized
		return frappe.get_doc("Smart Data Import", self.name)

	@frappe.whitelist()
	def start_import(self):
		return start_smart_import(self.name)

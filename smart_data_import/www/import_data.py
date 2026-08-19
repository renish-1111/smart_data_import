# Copyright (c) 2026, ERPNext AI Team and contributors
# For license information, please see license.txt

import frappe
import frappe.sessions  # frappe.sessions is not an attribute of frappe until imported

no_cache = 1


def get_context(context):
	# Guest could theoretically upload+import too if left open, but the engine writes
	# with ignore_permissions=True — anonymous access here is not an option.
	if frappe.session.user == "Guest":
		frappe.local.flags.redirect_location = "/login?redirect-to=/import_data"
		raise frappe.Redirect

	context.csrf_token = frappe.sessions.get_csrf_token()
	return context

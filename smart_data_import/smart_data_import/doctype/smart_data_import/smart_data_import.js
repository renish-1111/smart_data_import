// Copyright (c) 2026, ERPNext AI Team and contributors
// For license information, please see license.txt

frappe.ui.form.on("Smart Data Import", {
	refresh(frm) {
		// Auto-analyze files on refresh if pending or missing dependency graph
		if (frm.doc.docstatus === 0 && frm.doc.files && frm.doc.files.length > 0) {
			if (frm.doc.status === "Pending" || !frm.doc.dependencies || frm.doc.dependencies.length === 0) {
				frm.events.analyze_dependencies(frm);
			}
		}

		// Prominent 1-Click Primary Action Button
		if (frm.doc.docstatus === 0 && frm.doc.status !== "Processing" && frm.doc.status !== "Completed" && frm.doc.status !== "Partial Success") {
			frm.add_custom_button(__("🚀 Start Import Now"), () => {
				frm.events.start_import(frm);
			}).addClass("btn-primary btn-lg");

			frm.add_custom_button(__("🔄 Re-Analyze Files"), () => {
				frm.events.analyze_dependencies(frm);
			});
		} else {
			frm.clear_custom_buttons();
		}

		// Realtime Socket Progress Listener
		frappe.realtime.on("smart_import_progress", (data) => {
			if (data.doc_name === frm.doc.name) {
				frm.set_value("status", data.status);
				frm.set_value("imported_records", data.imported);
				frm.set_value("failed_records", data.failed);
				frm.set_value("progress_percent", data.progress);
				frm.refresh_fields();

				if (data.status === "Completed" || data.status === "Partial Success" || data.status === "Processing") {
					frm.clear_custom_buttons();
				}

				if (data.message) {
					frappe.show_progress(__("Import Progress"), data.progress, 100, data.message);
				}
			}
		});

		// Visual status banner alerts
		if (frm.doc.status === "Ready" && frm.doc.dependencies && frm.doc.dependencies.length > 0) {
			const sorted_deps = [...frm.doc.dependencies].sort((a, b) => a.execution_tier - b.execution_tier);
			const flow = sorted_deps.map(d => `${d.doctype_name} (${d.total_count} rows)`).join(" ➔ ");
			frm.dashboard.set_headline_alert(
				__("✨ Auto-Analyzed Execution Flow ({0} Files | {1} Total Rows): <b>{2}</b>. Ready to import!", 
					[frm.doc.files.length, frm.doc.total_records || 0, flow]),
				"green"
			);
		} else if (frm.doc.status === "Completed") {
			frm.dashboard.set_headline_alert(__("🎉 Data import completed successfully! {0} records inserted.", [frm.doc.imported_records]), "green");
		} else if (frm.doc.status === "Processing") {
			frm.dashboard.set_headline_alert(__("⏳ Processing data in high-speed background batches...", []), "blue");
		} else if (frm.doc.status === "Partial Success") {
			frm.dashboard.set_headline_alert(__("⚠️ Import finished with some errors. Download the failed rows log below."), "orange");
		}
	},

	after_save(frm) {
		// Re-analyze whenever files are saved, using reload_doc to prevent timestamp conflicts
		if (frm.doc.docstatus === 0 && frm.doc.files && frm.doc.files.length > 0 && !frm._is_analyzing) {
			frm._is_analyzing = true;
			frm.events.analyze_dependencies(frm);
		}
	},

	analyze_dependencies(frm) {
		if (!frm.doc.files || frm.doc.files.length === 0) return;

		frappe.call({
			method: "analyze_dependencies",
			doc: frm.doc,
			freeze: false,
			callback(r) {
				frm._is_analyzing = false;
				if (!r.exc) {
					// Synchronize document and timestamps from database
					frm.reload_doc();
				}
			}
		});
	},

	start_import(frm) {
		if (!frm.doc.files || frm.doc.files.length === 0) {
			frappe.msgprint(__("Please attach your Excel (.xlsx) or CSV data files first."));
			return;
		}

		frappe.confirm(
			__("Start high-speed batch import for all {0} attached files?", [frm.doc.files.length]),
			() => {
				frappe.call({
					method: "start_import",
					doc: frm.doc,
					freeze: true,
					freeze_message: __("Starting background import job..."),
					callback(r) {
						if (!r.exc) {
							frm.set_value("status", "Processing");
							frm.refresh();
							frappe.show_alert({
								message: __("Import job started! Tracking progress in real time."),
								indicator: "blue"
							});
						}
					}
				});
			}
		);
	}
});

// Reset status when files table changes
frappe.ui.form.on("Smart Data Import File", {
	files_add(frm) {
		frm.set_value("status", "Pending");
	},
	files_remove(frm) {
		frm.set_value("status", "Pending");
	},
	file(frm) {
		frm.set_value("status", "Pending");
	},
	doctype_name(frm) {
		frm.set_value("status", "Pending");
	}
});

// Copyright (c) 2026, ERPNext AI Team and contributors
// For license information, please see license.txt

const TERMINAL_STATUSES = ["Completed", "Failed", "Partial Success", "Rolled Back", "Cancelled"];
const BUSY_STATUSES = ["Processing", "Rolling Back"];

frappe.ui.form.on("Smart Data Import", {
	refresh(frm) {
		frm.events.bind_realtime(frm);
		frm.events.setup_buttons(frm);
		frm.events.render_headline(frm);
		frm.events.render_indicators(frm);

		// Analyze once per document, silently, when files were added but never analyzed.
		// The flag is essential: a file that yields no rows keeps the status at
		// "Pending", which would otherwise re-trigger analysis on every refresh.
		if (
			!frm.is_new() &&
			!frm.__analyzing &&
			frm.__auto_analyzed !== frm.doc.name &&
			(frm.doc.files || []).length &&
			frm.doc.status === "Pending" &&
			!(frm.doc.dependencies || []).length
		) {
			frm.__auto_analyzed = frm.doc.name;
			frm.events.analyze_dependencies(frm, true);
		}
	},

	after_save(frm) {
		if ((frm.doc.files || []).length && frm.doc.status === "Pending" && !frm.__analyzing) {
			frm.events.analyze_dependencies(frm, true);
		}
	},

	// ------------------------------------------------------------------ buttons

	setup_buttons(frm) {
		frm.clear_custom_buttons();
		if (frm.is_new()) {
			return;
		}

		const running = BUSY_STATUSES.includes(frm.doc.status);
		const finished = TERMINAL_STATUSES.includes(frm.doc.status);

		if (!running && !finished) {
			// Custom button (not set_primary_action) so the standard Save action keeps working.
			frm.add_custom_button(__("Start Import"), () => frm.events.start_import(frm)).addClass(
				"btn-primary"
			);
		}

		if (frm.doc.status === "Processing") {
			frm.add_custom_button(__("Cancel Import"), () => frm.events.cancel_import(frm)).addClass(
				"btn-danger"
			);
		}

		if ((frm.doc.errors || []).length) {
			frm.add_custom_button(__("View Errors"), () => frm.events.show_errors(frm)).addClass(
				"btn-warning"
			);
		}

		if (finished) {
			frm.add_custom_button(__("Reset for New Run"), () => frm.events.reset_import(frm));
			if (frm.doc.error_file) {
				frm.add_custom_button(__("Download Import Log"), () => {
					window.open(frm.doc.error_file);
				});
			}
		}

		// Undo is offered whenever a manifest of created records still exists.
		if (!running && frm.doc.rollback_file) {
			frm.add_custom_button(__("Rollback Import"), () => frm.events.rollback_import(frm)).addClass(
				"btn-danger"
			);
		}

		if (!running) {
			frm.add_custom_button(__("Preview & Column Mapping"), () => frm.events.preview_mapping(frm));
			// Same underlying analyze call the self-service /import_data portal calls
			// "Dry Run" — validates everything (detection, mapping, dependency order)
			// without writing any records, so the label is kept consistent across both.
			frm.add_custom_button(__("Dry Run"), () => frm.events.analyze_dependencies(frm));
			frm.add_custom_button(__("Download Template"), () => frm.events.download_template(frm));
		}
	},

	// ------------------------------------------------------------------ headline

	render_headline(frm) {
		if (frm.is_new()) {
			frm.dashboard.set_headline_alert(
				__(
					"<b>3 steps:</b> 1. <b>Download Template</b> for the DocType you want to import &nbsp;→&nbsp; 2. attach the filled file below &nbsp;→&nbsp; 3. click <b>Start Import</b>. Column order does not matter."
				),
				"blue"
			);
			return;
		}

		const warnings = (frm.doc.files || []).filter((r) => (r.mapping_summary || "").includes("⚠"));
		const problems = (frm.doc.files || []).filter((r) => r.status === "Failed" && !r.total_rows);

		if (frm.doc.status === "Processing") {
			frm.dashboard.set_headline_alert(
				__("Importing in the background — you can close this page, progress is saved."),
				"blue"
			);
		} else if (frm.doc.status === "Rolling Back") {
			frm.dashboard.set_headline_alert(
				__("Rolling back — deleting the records this import created..."),
				"orange"
			);
		} else if (frm.doc.status === "Rolled Back") {
			frm.dashboard.set_headline_alert(
				__("Rolled back: {0} records deleted.{1}", [
					frm.doc.rolled_back_records || 0,
					frm.doc.rollback_file
						? __(" Some records could not be deleted — see the Detailed Error Log.")
						: "",
				]),
				frm.doc.rollback_file ? "orange" : "green"
			);
		} else if (frm.doc.status === "Cancelled") {
			frm.dashboard.set_headline_alert(
				__("Cancelled: {0} records were imported before it was stopped.", [
					frm.doc.imported_records || 0,
				]),
				"orange"
			);
		} else if (frm.doc.status === "Completed") {
			frm.dashboard.set_headline_alert(
				__("Done: {0} records imported in {1}s ({2} skipped).", [
					frm.doc.imported_records || 0,
					frm.doc.execution_time_seconds || 0,
					frm.doc.skipped_records || 0,
				]),
				"green"
			);
		} else if (frm.doc.status === "Partial Success") {
			frm.dashboard.set_headline_alert(
				__("{0} imported, {1} failed, {2} skipped. {3}", [
					frm.doc.imported_records || 0,
					frm.doc.failed_records || 0,
					frm.doc.skipped_records || 0,
					error_link(__("See what went wrong")),
				]),
				"orange"
			);
			bind_error_link(frm);
		} else if (frm.doc.status === "Failed") {
			frm.dashboard.set_headline_alert(
				__("Import failed. {0}, fix the cause, then Reset for New Run.", [
					error_link(__("See what went wrong")),
				]),
				"red"
			);
			bind_error_link(frm);
		} else if (problems.length) {
			frm.dashboard.set_headline_alert(
				__("{0} file(s) cannot be imported yet — see the Result / Error Log column.", [
					problems.length,
				]),
				"red"
			);
		} else if (frm.doc.status === "Ready" && (frm.doc.dependencies || []).length) {
			const flow = [...frm.doc.dependencies]
				.sort((a, b) => a.execution_tier - b.execution_tier)
				.map((d) => `${d.doctype_name} (${d.total_count})`)
				.join(" → ");
			const suffix = warnings.length
				? __(" — {0} file(s) have mapping warnings, use Preview first.", [warnings.length])
				: __(" — ready to import.");
			frm.dashboard.set_headline_alert(
				__("{0} rows in this order: <b>{1}</b>{2}", [frm.doc.total_records || 0, flow, suffix]),
				warnings.length ? "orange" : "green"
			);
		}
	},

	// ------------------------------------------------------------------ indicators

	render_indicators(frm) {
		if (frm.is_new()) {
			return;
		}
		if (frm.doc.imported_records) {
			frm.dashboard.add_indicator(__("{0} imported", [frm.doc.imported_records]), "green");
		}
		if (frm.doc.failed_records) {
			frm.dashboard.add_indicator(__("{0} failed", [frm.doc.failed_records]), "red");
		}
		if (frm.doc.skipped_records) {
			frm.dashboard.add_indicator(__("{0} skipped", [frm.doc.skipped_records]), "orange");
		}
		if (frm.doc.rolled_back_records) {
			frm.dashboard.add_indicator(__("{0} rolled back", [frm.doc.rolled_back_records]), "purple");
		}
	},

	// ------------------------------------------------------------------ realtime

	bind_realtime(frm) {
		if (frm.__sdi_realtime_bound) {
			return;
		}
		frm.__sdi_realtime_bound = true;

		frappe.realtime.on("smart_import_progress", (data) => {
			if (!frm.doc || data.doc_name !== frm.doc.name) {
				return;
			}

			// Update in place: never use set_value here, it would mark the form dirty.
			frm.doc.status = data.status;
			frm.doc.imported_records = data.imported;
			frm.doc.failed_records = data.failed;
			frm.doc.skipped_records = data.skipped || 0;
			frm.doc.rolled_back_records = data.rolled_back || 0;
			frm.doc.progress_percent = data.progress;
			frm.refresh_fields();

			const progress_title =
				data.status === "Rolling Back"
					? __("Rollback Progress")
					: data.status === "Analyzing"
					  ? __("Dry Run Progress")
					  : __("Import Progress");
			frm.dashboard.show_progress(progress_title, data.progress || 0, data.message || __("Working..."));

			if (TERMINAL_STATUSES.includes(data.status)) {
				frm.dashboard.hide_progress();
				frappe.show_alert({
					message: data.message || __("Import finished."),
					indicator: data.status === "Completed" ? "green" : "orange",
				});
				const had_problems = (data.failed || 0) > 0 || (data.skipped || 0) > 0;
				Promise.resolve(frm.reload_doc()).then(() => {
					// Show the reasons straight away instead of making the user hunt for them.
					if (had_problems && (frm.doc.errors || []).length) {
						frm.events.show_errors(frm);
					}
				});
			}
		});
	},

	// ------------------------------------------------------------------ actions

	analyze_dependencies(frm, silent) {
		if (!(frm.doc.files || []).length || frm.__analyzing) {
			return;
		}
		frm.__analyzing = true;
		// A live progress bar (fed by the "smart_import_progress" realtime event) reads
		// better here than a blocking freeze overlay, which would just hide it.
		frm.dashboard.show_progress(__("Dry Run Progress"), 0, __("Reading your files..."));

		frm.call({
			method: "analyze_dependencies",
			doc: frm.doc,
		})
			.then(() => frm.reload_doc())
			.always(() => {
				frm.dashboard.hide_progress();
				frm.__analyzing = false;
			});
	},

	start_import(frm) {
		if (!(frm.doc.files || []).length) {
			frappe.msgprint(__("Attach at least one Excel (.xlsx) or CSV file first."));
			return;
		}
		if (frm.is_dirty()) {
			frappe.msgprint(__("Please save the document before starting the import."));
			return;
		}

		const rows = frm.doc.total_records || 0;
		const mode = frm.doc.import_type || __("Insert New Records");
		frappe.confirm(
			__("Import {0} rows from {1} file(s) using mode <b>{2}</b>?", [
				rows,
				frm.doc.files.length,
				mode,
			]),
			() => {
				frm.call({
					method: "start_import",
					doc: frm.doc,
					freeze: true,
					freeze_message: __("Queueing background import..."),
				}).then((r) => {
					const result = r.message || {};
					if (result.status === "error") {
						frappe.msgprint(result.message);
						return;
					}
					frappe.show_alert({ message: result.message, indicator: "blue" });
					frm.reload_doc();
				});
			}
		);
	},

	show_errors(frm) {
		frm.call({
			method: "error_summary",
			doc: frm.doc,
			args: { limit: 25 },
			freeze: true,
			freeze_message: __("Grouping the errors..."),
		}).then((r) => {
			const data = r.message;
			if (!data || !(data.groups || []).length) {
				frappe.msgprint(__("No row errors were recorded for this import."));
				return;
			}

			const dialog = new frappe.ui.Dialog({
				title: __("Import Errors"),
				size: "large",
				fields: [{ fieldname: "errors", fieldtype: "HTML" }],
				primary_action_label: data.log_file ? __("Download Full Log") : __("Close"),
				primary_action() {
					if (data.log_file) {
						window.open(data.log_file);
					}
					dialog.hide();
				},
			});
			dialog.fields_dict.errors.$wrapper.html(render_errors(data));
			dialog.show();
		});
	},

	rollback_import(frm) {
		frm.call({
			method: "rollback_summary",
			doc: frm.doc,
			freeze: true,
			freeze_message: __("Reading the rollback manifest..."),
		}).then((r) => {
			const info = r.message;
			if (!info) return;

			if (!info.deletable) {
				frappe.msgprint({
					title: __("Nothing to roll back"),
					indicator: "orange",
					message: info.updated
						? __(
								"This import updated {0} existing record(s). Updates cannot be undone automatically — restore them from each document's version history.",
								[info.updated]
						  )
						: __("No record of created documents exists for this import."),
				});
				return;
			}

			const rows = (info.doctypes || [])
				.map(
					(d) =>
						`<tr><td>${frappe.utils.escape_html(d.doctype)}</td>
						<td class="text-right">${d.inserted}</td>
						<td class="text-right text-muted">${d.updated || 0}</td></tr>`
				)
				.join("");

			const dialog = new frappe.ui.Dialog({
				title: __("Rollback Import"),
				fields: [
					{ fieldname: "summary", fieldtype: "HTML" },
					{
						fieldname: "force",
						fieldtype: "Check",
						label: __("Ignore link validations (dangerous)"),
						default: 0,
						description: __(
							"Only tick this if deletion fails because other documents link to these records. It deletes them anyway and can leave broken links behind."
						),
					},
				],
				primary_action_label: __("Delete {0} records", [info.deletable]),
				primary_action(values) {
					dialog.hide();
					frappe.confirm(
						__("Permanently delete {0} records created by this import?", [info.deletable]),
						() => {
							frm.call({
								method: "rollback",
								doc: frm.doc,
								args: { force: values.force ? 1 : 0 },
								freeze: true,
								freeze_message: __("Queueing rollback..."),
							}).then((res) => {
								const out = res.message || {};
								if (out.status === "error") {
									frappe.msgprint(out.message);
									return;
								}
								frappe.show_alert({ message: out.message, indicator: "orange" });
								frm.reload_doc();
							});
						}
					);
				},
			});

			dialog.fields_dict.summary.$wrapper.html(`
				<div class="alert alert-warning" style="padding:8px 12px;">
					${__("This deletes the records this import created. It cannot be undone.")}
					${
						info.updated
							? "<br>" +
							  __("{0} updated record(s) will be left untouched — updates cannot be reverted.", [
									info.updated,
							  ])
							: ""
					}
				</div>
				<table class="table table-bordered table-sm" style="font-size:12px;">
					<thead><tr>
						<th>${__("DocType")}</th>
						<th class="text-right">${__("To delete")}</th>
						<th class="text-right">${__("Updated (kept)")}</th>
					</tr></thead>
					<tbody>${rows}</tbody>
				</table>
				<div class="text-muted small">${__("Records are deleted newest first, so children go before their parents.")}</div>
			`);
			dialog.show();
		});
	},

	cancel_import(frm) {
		frappe.confirm(
			__("Stop this import? Records already imported so far are kept — this cannot be undone by itself, but you can Rollback afterward if needed."),
			() => {
				frm.call({ method: "cancel_import", doc: frm.doc, freeze: true }).then((r) => {
					const result = r.message || {};
					if (result.status === "error") {
						frappe.msgprint(result.message);
						return;
					}
					frappe.show_alert({ message: result.message, indicator: "orange" });
					frm.reload_doc();
				});
			}
		);
	},

	reset_import(frm) {
		frappe.confirm(
			__("Clear all counters and logs so this import can run again? Imported records are not deleted."),
			() => {
				frm.call({ method: "reset_for_rerun", doc: frm.doc, freeze: true }).then(() =>
					frm.reload_doc()
				);
			}
		);
	},

	download_template(frm) {
		const preset = [...new Set((frm.doc.files || []).map((r) => r.doctype_name).filter(Boolean))];
		const SEPARATE_LABEL = __("Separate files (.zip)");
		const COMBINED_LABEL = __("Combined into one file (one sheet each)");

		const dialog = new frappe.ui.Dialog({
			title: __("Download Import Template"),
			fields: [
				{
					fieldname: "target_doctypes",
					fieldtype: "MultiSelectList",
					label: __("DocType(s)"),
					reqd: 1,
					default: preset,
					get_data: function (txt) {
						return frappe.db.get_link_options("DocType", txt, { istable: 0, issingle: 0 });
					},
				},
				{
					fieldname: "include_optional",
					fieldtype: "Check",
					label: __("Include optional fields"),
					default: 1,
					description: __(
						"Unchecked gives you only the mandatory columns — the fastest way to a valid file."
					),
				},
				{
					fieldname: "package_mode",
					fieldtype: "Select",
					label: __("When more than one DocType is selected"),
					options: [COMBINED_LABEL, SEPARATE_LABEL].join("\n"),
					default: COMBINED_LABEL,
					description: __(
						"Combined puts one sheet per DocType in a single workbook. Separate gives you one standalone .xlsx per DocType, zipped together."
					),
				},
			],
			primary_action_label: __("Generate"),
			primary_action(values) {
				const doctypes = values.target_doctypes || [];
				if (!doctypes.length) {
					frappe.msgprint(__("Select at least one DocType."));
					return;
				}
				frm.call({
					method: "get_template",
					doc: frm.doc,
					args: {
						target_doctype: doctypes,
						include_optional: values.include_optional,
						mode: values.package_mode === SEPARATE_LABEL ? "separate" : "single",
					},
					freeze: true,
					freeze_message: __("Building template(s)..."),
				}).then((r) => {
					dialog.hide();
					const info = r.message;
					if (!info) return;
					window.open(info.file_url);
					const summary = (info.stats || [])
						.map((s) => `${s.doctype} (${s.total_columns}, ${s.mandatory_columns} required)`)
						.join(", ");
					frappe.show_alert({
						message: summary ? __("Downloaded: {0}", [summary]) : __("Template downloaded."),
						indicator: "green",
					});
				});
			},
		});
		dialog.show();
	},

	preview_mapping(frm) {
		if (!(frm.doc.files || []).length) {
			frappe.msgprint(__("Attach a file first."));
			return;
		}

		frm.call({
			method: "preview_mapping",
			doc: frm.doc,
			args: { sample_size: 5 },
			freeze: true,
			freeze_message: __("Reading a few rows..."),
		}).then((r) => {
			const data = r.message;
			if (!data) return;

			const skeleton = {};
			(data.files || []).forEach((f) => {
				if (f.doctype && (f.unmapped || []).length) {
					skeleton[f.doctype] = skeleton[f.doctype] || { column_map: {} };
					f.unmapped.forEach((h) => {
						skeleton[f.doctype].column_map[h] = "";
					});
				}
			});

			const dialog = new frappe.ui.Dialog({
				title: __("Preview & Column Mapping"),
				size: "extra-large",
				fields: [{ fieldname: "preview", fieldtype: "HTML" }],
				primary_action_label: Object.keys(skeleton).length
					? __("Add Mapping Skeleton")
					: __("Close"),
				primary_action() {
					if (Object.keys(skeleton).length) {
						const existing = frm.doc.filter_rules_json
							? JSON.parse(frm.doc.filter_rules_json)
							: {};
						frm.set_value(
							"filter_rules_json",
							JSON.stringify(Object.assign(existing, skeleton), null, 2)
						);
						frappe.show_alert({
							message: __("Skeleton added — fill in the fieldnames and save."),
							indicator: "blue",
						});
					}
					dialog.hide();
				},
			});
			dialog.fields_dict.preview.$wrapper.html(render_preview(data));
			dialog.show();
		});
	},
});

function error_link(label) {
	return `<a class="sdi-show-errors" style="text-decoration:underline;cursor:pointer;">${label}</a>`;
}

function bind_error_link(frm) {
	// The headline lives in the layout message area and is re-rendered on every
	// refresh, so re-bind each time.
	const $message = frm.layout && frm.layout.message;
	if (!$message || !$message.length) {
		return;
	}
	$message
		.find(".sdi-show-errors")
		.off("click.sdi")
		.on("click.sdi", () => frm.events.show_errors(frm));
}

function render_errors(data) {
	const esc = frappe.utils.escape_html;
	const badge = {
		Failed: "red",
		Skipped: "orange",
		Hierarchy: "blue",
		Rollback: "purple",
	};

	let html = `<div class="mb-3">
		<span class="indicator-pill red">${__("{0} failed", [data.failed || 0])}</span>
		<span class="indicator-pill orange">${__("{0} skipped", [data.skipped || 0])}</span>
	</div>
	<div class="text-muted small mb-2">${__("Grouped by cause, most frequent first.")}</div>
	<table class="table table-bordered table-sm" style="font-size:12px;">
		<thead><tr>
			<th style="width:60px;" class="text-right">${__("Rows")}</th>
			<th style="width:90px;">${__("Type")}</th>
			<th>${__("Problem")}</th>
			<th style="width:120px;">${__("DocType")}</th>
			<th style="width:70px;" class="text-right">${__("e.g. row")}</th>
		</tr></thead><tbody>`;

	(data.groups || []).forEach((g) => {
		html += `<tr>
			<td class="text-right"><b>${g.count}</b></td>
			<td><span class="indicator-pill ${badge[g.type] || "gray"}">${esc(g.type)}</span></td>
			<td>${esc(g.problem)}<div class="text-muted" style="font-size:11px;">${esc(
				g.example_reason || ""
			)}</div></td>
			<td>${esc(g.doctypes || "")}</td>
			<td class="text-right">${g.example_row || ""}</td>
		</tr>`;
	});

	html += `</tbody></table>`;

	if (data.truncated) {
		html += `<div class="text-muted small">${__(
			"Only the first {0} problem rows are kept on this document — download the full log for the rest.",
			[data.shown_rows]
		)}</div>`;
	}
	html += `<div class="text-muted small mt-2">${__(
		"Every problem row is also listed under <b>Failed & Skipped Rows</b> on the form."
	)}</div>`;
	return html;
}

function render_preview(data) {
	const esc = frappe.utils.escape_html;
	let html = `<div class="text-muted small mb-3">${__("Import mode")}: <b>${esc(
		data.import_type || ""
	)}</b></div>`;

	(data.files || []).forEach((f) => {
		html += `<div style="margin-bottom:24px;">
			<h5>${__("File {0}", [f.idx])}: ${esc(f.file || "")} ${
				f.sheet ? `<span class="text-muted small">(${esc(f.sheet)})</span>` : ""
			}</h5>`;

		if (f.error) {
			html += `<div class="alert alert-danger">${esc(f.error)}</div></div>`;
			return;
		}

		html += `<div class="mb-2">${__("Target")}: <b>${esc(f.doctype)}</b> &middot; ${__(
			"{0} data rows",
			[f.total_rows]
		)}</div>`;

		if ((f.missing_mandatory || []).length) {
			html += `<div class="alert alert-warning" style="padding:6px 10px;">⚠ ${__(
				"Mandatory fields missing from this file"
			)}: <b>${f.missing_mandatory.map(esc).join(", ")}</b>. ${__(
				"Add the columns, set them under Defaults, or tick 'Ignore Missing Mandatory Fields'."
			)}</div>`;
		}
		if ((f.unmapped || []).length) {
			html += `<div class="alert alert-warning" style="padding:6px 10px;">${__(
				"These columns will be ignored"
			)}: <b>${f.unmapped.map(esc).join(", ")}</b></div>`;
		}

		const mandatory_set = new Set(f.mandatory_fieldnames || []);
		const mandatory_style = 'style="color:#c0392b;font-weight:600;"';

		const mapped_pairs = Object.entries(f.mapped || {});
		html += `<table class="table table-bordered table-sm" style="font-size:12px;">
			<thead><tr>
				<th>${__("Column in file")}</th><th>${__("Imported into field")}</th>
			</tr></thead><tbody>`;
		mapped_pairs.forEach(([header, fieldname]) => {
			const is_mandatory = mandatory_set.has(fieldname);
			html += `<tr><td>${esc(header)}</td><td><code${is_mandatory ? ` ${mandatory_style}` : ""}>${esc(
				fieldname
			)}${is_mandatory ? " *" : ""}</code></td></tr>`;
		});
		(f.unmapped || []).forEach((header) => {
			html += `<tr class="text-muted"><td>${esc(header)}</td><td>— ${__("ignored")}</td></tr>`;
		});
		html += `</tbody></table>`;
		if (mandatory_set.size) {
			html += `<div class="text-muted small mb-2"><span ${mandatory_style}>*</span> ${__(
				"Mandatory field on the target DocType."
			)}</div>`;
		}

		if ((f.samples || []).length) {
			html += `<div class="text-muted small mb-1">${__("First rows as they will be read")}:</div>
				<div style="overflow-x:auto;"><table class="table table-bordered table-sm" style="font-size:11px;"><thead><tr>`;
			(f.headers || []).forEach((h, i) => {
				const target = (f.mapped || {})[String(h).trim()];
				const is_mandatory = target && mandatory_set.has(target);
				const th_style = is_mandatory ? ` ${mandatory_style}` : !target ? ' class="text-muted"' : "";
				html += `<th${th_style}>${esc(h || `col ${i + 1}`)}</th>`;
			});
			html += `</tr></thead><tbody>`;
			f.samples.forEach((row) => {
				html += "<tr>" + row.map((v) => `<td>${esc(v)}</td>`).join("") + "</tr>";
			});
			html += `</tbody></table></div>`;
		}
		html += `</div>`;
	});

	return html;
}

// Any change to the file list invalidates the previous analysis.
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
	},
	sheet_name(frm) {
		frm.set_value("status", "Pending");
	},
});

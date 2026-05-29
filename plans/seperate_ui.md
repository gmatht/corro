Separate UI and Core: Plan

Goal
- Keep UI-specific behaviour, heuristics, and transient maintenance operations out of the persistent core and append-only log.
- Make clear boundaries so the core (ops, grid, io) has a small, well-tested API and the UI continues to own UX heuristics and rendering-only changes.

Principles
- Minimal invasive changes: gate persistence at the UI caller sites rather than changing core append/parse semantics.
- Preserve user intent: user-initiated commands (explicit menu choices, Enter-confirmed values, SetColWidth/SetMaxColWidth modes, sheet creation/rename/move) remain persistent.
- Treat layout/render maintenance (auto-fit, auto-grow, synthetic SIZE/COL_WIDTH/MAX_COL_WIDTH updates, and similar transient adjustments) as in-memory-only unless the user explicitly requests persistence.
- Keep undo/redo user-focused: the persistent undo history should reflect user actions; synthetic housekeeping should not pollute the persistent log or permanent undo history.

Scope
- Files/areas primarily involved:
  - src/ui/mod.rs (caller sites, navigation, edit flows, apply paths)
  - src/io/mod.rs (commit_workbook_op, commit_workbook_set_column_format_batch, tail_apply_workbook)
  - src/ops/mod.rs (Op variants and to_log_lines_with_policy)
  - src/grid/mod.rs (auto-fit, set_col_width, grow_main_* methods)
- Tests touching unsaved-file creation, commit_workbook_op, and save_workbook snapshoting.

High-level plan (phases)

Phase 1 — Audit and classification (read-only)
- Enumerate all commit_workbook_op and commit_workbook_set_column_format_batch call-sites and classify each as:
  - user-initiated (persist), or
  - UI-maintenance / synthetic (do not persist)
- Mark ambiguous sites for review.
- Produce a short per-site change-list that specifies whether to (A) leave, (B) replace commit_workbook_op with in-memory apply, or (C) refactor into a new ui helper.

Phase 2 — Small, local caller-side API & helper
- Add a single thin helper in the UI module (or a small ui::persistence helper):
  - apply_user_op(op) — pushes inverse op to undo history and uses commit_workbook_op when path present.
  - apply_ui_maint_op(op) — applies op in-memory only, does not append to log, and does NOT push to persistent undo history (optionally push transient undo if UI requires immediate revert during session).
- Replace only the classified synthetic call-sites with apply_ui_maint_op(op).
- Replace explicit user commands to use apply_user_op(op) (where not already centralized).
- Keep commit_workbook_op and io::append behavior unchanged.

Phase 3 — Tests and validation
- Add focused unit/integration tests:
  - confirm unsaved-file commit (auto-create) still writes CORRO_LOG header and user op lines.
  - confirm that running UI maintenance paths (auto-fit, bulk layout) do NOT append SIZE/COL_WIDTH/MAX_COL_WIDTH lines to an on-disk unsaved file created by the UI.
  - regression test for the bug: "typing into a field and immediately moving Right into a blank column should move to the next main column (not margin)" — reproduce, assert cursor lands on next main column and edit_target follows.
  - preserve existing behavior for Mode::SetColWidth / SetMaxColWidth and other explicit user modes.

Phase 4 — Rollout and cleanup
- Merge small changes behind a feature branch.
- Run full test suite and a set of manual UI scenarios.
- Update DESIGN.md / README to document where persistence decisions are made.
- Optionally, consolidate repetitive patterns (many call sites follow if self.path.is_some() { commit } else { op.apply }) into the helper created in Phase 2.

Concrete tasks (practical checklist)
1. Repo scan: list every commit_workbook_op and commit_workbook_set_column_format_batch call-site and create classification file (CSV or table) for quick review.
2. Create UI helper(s): apply_user_op, apply_ui_maint_op (small, documented functions in src/ui/mod.rs).
3. For each synthetic site, replace the commit_workbook_op call with apply_ui_maint_op and remove persistent undo push where appropriate.
4. For user-initiated places that already call commit_workbook_op, ensure they push inverse op to op_history before commit (preserve existing semantics).
5. Add tests described in Phase 3.
6. Run test suite, fix failures, and iterate.

Addressing the "edit then Right into blank column lands in margin" bug
- Reproduce as a unit/regression test describing the exact sequence:
  1) Start with main columns N.
 2) Enter edit mode in a main cell at the rightmost main column.
 3) Commit the edit and immediately press Right (or programmatically call the function that performs the movement).
 4) Assert the cursor moves into the next main column (growing main area when necessary) rather than into the right-margin.
- Likely fixes to try (minimal first):
  - Ensure move_cursor_one_col_horizontal grows main cols before computing the new logical cell address and clamps — this ensures the "next" main column exists.
  - Ensure edit_target_addr / cursor sync logic (maybe_sync_edit_target_with_highlighted_cell and commit_edit_and_move_down/right helpers) updates the cursor and edit_target consistently when the user moves while in or immediately after edit mode.
  - Add test and iterate until behavior matches expectation.

API/Refactor suggestions (low risk)
- Do not change io/commit_workbook_op semantics in the first pass. Gate persistence at caller level.
- Add helper(s) in UI with a clear name and doc comment so future contributors know which to call for persistent vs transient ops.
- If multiple call-sites converge to the same pattern, centralize them behind the helper.

Acceptance criteria
- No synthetic layout ops are appended to the on-disk unsaved file in new tests.
- Existing user-visible persistent operations behave unchanged and tests continue to pass.
- The specific edit->Right bug is covered by a regression test and fixed.

Risks & mitigations
- Risk: unintentionally drop user-intended persistence. Mitigation: conservative classification; default to persisting if unclear.
- Risk: undo/redo divergence between UI and persisted log. Mitigation: centralize undo push in apply_user_op and document semantics.

Next actions (first PR)
1. Produce the audit table (Phase 1) as a separate artifact/PR for review.
2. Implement the UI helper and change 3–5 obvious synthetic call-sites as a small PR with tests.
3. Iterate on the full conversion once reviewed.

Notes
- The repository already contains some helper logic and guard comments (save_to_path detects synthetic ops and omits them on final save). This plan follows the same safety-first approach but makes the caller-level intent explicit and auditable.

Appendix: Per-Call-Site Audit Summary

I audited the repository for direct call-sites of `commit_workbook_op` and `commit_workbook_set_column_format_batch` (primarily in `src/ui/mod.rs`). Below are concise per-call-site entries with a recommendation whether the call-site should persist the op (call `commit_workbook_op`) or keep the change UI-only / in-memory (use an in-memory apply helper). For ambiguous cases I mark them as `Ambiguous` and give a short rationale.

Format: <file:path> — <approx line> — <enclosing function / context> — <emitted WorkbookOp/Op> — <recommended classification> — <rationale / notes>

/root/src/corro/src/ui/mod.rs — 2920 — edit commit path (generic cell commit) — WorkbookOp::SheetOp(Op::SetCell / Op::SetCellRef / SetMainSize) — Persist — This is user-entered cell content or explicit size; preserve in log.

/root/src/corro/src/ui/mod.rs — 2947 — commit after edit with fit-to-content — WorkbookOp::SheetOp(Op::SetCell + auto-fit SIZE/COL_WIDTH) — Ambiguous — User committed a cell but UI also autotriggers fit; persist the SetCell but do not persist synthetic SIZE/COL_WIDTH unless user requested.

/root/src/corro/src/ui/mod.rs — 2970 — commit after paste/import flows — WorkbookOp::SheetOp(Op::FillRange / SetMainSize) — Persist — Paste/import is explicit user action; persistent.

/root/src/corro/src/ui/mod.rs — 3002 — commit in replace/atomic flows — WorkbookOp::SheetOp(Op::SetCell or FillRange) — Persist — User-driven content change; persist.

/root/src/corro/src/ui/mod.rs — 3138/3143 — format -> SetColumnFormat batch commit path — commit_workbook_set_column_format_batch (sheet_id, ops: &[Op::SetColumnFormat|SetAllColumnFormat]) — Persist — Explicit formatting chosen by the user; batching already consolidates validation and append, keep persisted.

/root/src/corro/src/ui/mod.rs — 4301/4304 — Cut operation commit path (menu Cut) — Op::FillRange (clearing cells) — Persist — User action; persist.

/root/src/corro/src/ui/mod.rs — 6117/6131 — Move rows commit (move_selected_rows_by_one / row movement) — WorkbookOp::SheetOp(Op::MoveRowRange / DeleteRowRange / DuplicateRow) — Persist — Explicit user row move/insert/delete; persist.

/root/src/corro/src/ui/mod.rs — 6177/6180/6202/6205/6247/6272/6275 — Column move/duplicate/delete flows — WorkbookOp::SheetOp(Op::MoveColRange / DuplicateColRange / DeleteColRange / SetColWidth / SetMaxColWidth) — Persist for user-invoked Move/Duplicate/Delete; SetColWidth/SetMaxColWidth persist only when invoked via explicit user Mode (SetColWidth / SetMaxColWidth). Otherwise, for UI auto-fit adjustments, do not persist. — Mixed

/root/src/corro/src/ui/mod.rs — 6921/6924/6946/6949/7023/7026 — Insert rows/columns & mitosis insert flows — WorkbookOp::SheetOp(Op::Insert-like: DuplicateRow / DuplicateCol / SetMainSize) — Persist — User-invoked structural changes; persist.

/root/src/corro/src/ui/mod.rs — 7580/7626 — Sort view / Save sort / SetViewSortCols — WorkbookOp::SheetOp(Op::SetViewSortCols) — Persist — Sorting persistence expected when user requests SaveSort; ephemeral sort views should not commit unless persist flag set.

/root/src/corro/src/ui/mod.rs — 10502 — Save / SaveAs flows that call commit_workbook_op — WorkbookOp variants (many) — Persist — These are explicit persistence points.

/root/src/corro/src/ui/mod.rs — 11614 / 11652 / 11678 / 11734 / 11921 / 11924 — Formatting menu flows (apply number/align/format to target) — WorkbookOp::SheetOp(Op::SetColumnFormat / SetAllColumnFormat / SetCellFormat) — Persist — User-chosen format changes should be persisted; for Scope::All the batch helper is preferred.

/root/src/corro/src/ui/mod.rs — 12786 / 12789 / 12816 / 12819 / 13050 / 13053 / 13105 / 13108 / 13166 / 13169 — Sheet-level operations (NewSheet / RenameSheet / CopySheet / MoveSheet / BalanceReport) — WorkbookOp::NewSheet / RenameSheet / CopySheet / MoveSheet / BalanceReport — Persist — Explicit user-level workbook changes.

/root/src/corro/src/ui/mod.rs — 13658 — external app.push_inverse_op + commit_workbook_op usage in tests / helpers — Persist — Test scaffolding and structured commits should persist.

Notes & guidance summary
- Default to persisting explicit user actions (menu commands, edit commits, structural changes, formatting when user confirms). These keep existing semantics and test expectations.
- Do not persist synthetic UI maintenance ops: auto-fit, transient column width updates, reactive SIZE changes produced solely for rendering. Where the caller currently appends such ops, prefer using an in-memory apply helper that updates the live workbook/grid and optionally pushes a transient inverse op for in-session undo.
- For ambiguous mixed flows (cell commit that triggers auto-fit), split the action: persist the explicit user op (SetCell/FillRange) and apply maintenance UI ops in-memory only.
- The `commit_workbook_set_column_format_batch` call-sites are already correct for persistence: they represent explicit user formatting operations and batch them for performance.

If you'd like, next I'll:
1) Commit this change (append-only update already staged in this patch). — DONE
2) Implement the small UI helper functions and convert a short list (3–5) of obvious synthetic call-sites to in-memory applies in a follow-up patch.

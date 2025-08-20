
# Kafaa PACE v23 Patch — True Edge, PACE-weighted Actions, Arabic i18n

This small patch adds three upgrades:
1) **Edge with percentiles** vs industry benchmarks (and optional history).
2) **Countermeasures** auto-prioritized by **PACE**.
3) **Arabic** labels on the PACE page (English default).

## Files
- `templates_add_i18n_and_edge.yaml` → append to `templates.yaml`.
- `engine_add_compute_edge.py` → paste into `engine.py` (new functions).
- `engine_edit_compute_pace.txt` → replace your existing `compute_pace(...)` with this version (signature changed).
- `engine_edit_propose_countermeasures.txt` → edit your `propose_countermeasures(...)` to accept `pace=...` and apply the weighting logic.
- `app_edits_i18n_pace_and_cm.txt` → add a language toggle, measured Edge inputs, and pass PACE to countermeasures.

## Notes
- **Measured values**: users can optionally enter FPY, changeover time, inventory days, loading time, FG aging. If left blank, the Edge effect defaults to neutral (1.0).
- **History**: you can later set `st.session_state["pace_history"] = {"fpy_pct":[95,97,99], "inventory_days":[60,45,30], ...}` to nudge Edge factors by historical percentiles.
- **Arabic**: only the PACE page is localized here; we can extend to more pages once you’re happy with the terms.
- The Edge factor is gently clamped to avoid overpowering money/severity; tune in `compute_edge_percentiles` if desired.

After applying, redeploy and test: PACE should show Edge multipliers working, and the Countermeasures page should reflect PACE priorities (more “Now” in the top themes).

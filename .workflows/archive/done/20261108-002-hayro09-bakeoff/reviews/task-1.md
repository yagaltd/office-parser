# Verdict — TASK 1 (scout): hayro 0.9.0 tree map

ok:true

Reviewer: orchestrator (gate), 2026-11-09.

Evidence: report persisted to .workflows/scout/20261108-002.md (14.7 KB);
precision audited before acceptance — all 10 families carry path:line
anchors in the 0.9.0 tree (spot-checked: context.rs MAX_NESTED_DEPTH,
hayro/src/lib.rs discarded group flags, cmap parse strictness sites,
page.rs OnceLock decode); glue section lists the exact new signatures
(render()/Pixmap/PixmapSettings, RenderCache lifetime, into_png) with the
~20-line glue shape; registry check proved the offline gap (8 crates need
one networked fetch or vendoring) and confirmed jbig2 0.4.0/postscript
0.2.0 fully satisfy the 0.9.0 consumers while our unwired vendored copies
are API-incompatible dead weight. Wave-1 briefs incorporate the map.

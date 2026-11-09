#!/usr/bin/env python3
"""pdfium vs hayro A/B bake-off runner (TASK 3; feeds the TASK 4 flip decision).

For each renderer, pinned via OFFICE_PARSER_RENDERER, it runs the
office-parser CLI over:

  (a) the synthetic corpus  dense_scan.pdf
      (--format json --json-no-image-bytes --ocr --ocr-force --split-pages;
      per-page transcripts from the JSONL -> WER against words.json)
  (b) tests/fixtures/pdf_*.pdf
      (plain --ocr so text-layer/OCR routing is respected; no WER -- no
      ground truth exists for fixtures)
  (c) real-scans/*.pdf  drop-in slot
      (--ocr --ocr-force; char/word delta ONLY, no WER column; the run
      succeeds when the directory is empty)

Metrics per renderer per corpus item: WER (synthetic only), char/word counts
(+ delta vs the other renderer), block count, low-confidence page count
(metadata.extra.ocr.low_confidence_pages), wall time. Output: report.md (one
markdown table per corpus section) + a stdout summary.

Arm integrity (AGENTS.md: never optional-but-silent): renderer selection in
src/formats/pdf.rs falls through SILENTLY (pdfium -> hayro -> pdftoppm), so
the runner (1) sniffs the CLI binary for hayro/pdfium markers, (2) mirrors
pdfium_lib_path(), and (3) hard-fails before measuring when a pinned backend
could not have served its arm. If both arms still produce byte-identical
synthetic transcripts, report.md carries a loud silent-fallback warning.

--selfcheck: generator determinism (two seeded runs -> sha256 equality) +
the WER scorer micro-case (wer.py --selfcheck). Exit 0 on pass.

python3 stdlib only. Nothing here is wired into CI or the test suite.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

HARNESS_DIR = Path(__file__).resolve().parent
REPO_ROOT = HARNESS_DIR.parents[2]
GENERATOR = HARNESS_DIR / "generate_corpus.py"
WER_PY = HARNESS_DIR / "wer.py"
DEFAULT_CLI = REPO_ROOT / "target" / "release" / "office-parser-cli"
DEFAULT_CORPUS_DIR = HARNESS_DIR
DEFAULT_FIXTURES_DIR = REPO_ROOT / "tests" / "fixtures"
DEFAULT_REAL_SCANS_DIR = HARNESS_DIR / "real-scans"
DEFAULT_WORK_DIR = Path(tempfile.gettempdir()) / "render-bakeoff"
DEFAULT_REPORT = HARNESS_DIR / "report.md"
RENDERERS = ("pdfium", "hayro")
FLIP_MARGIN = 0.01  # WER(hayro) <= WER(pdfium) + 1pt flips the default (TASK 4)

# Feature-gated string markers: present in the binary only when the
# corresponding render feature is compiled in. b"libpdfium.so" is NOT a
# usable pdfium marker -- install_pdfium() carries that literal regardless
# of features. b"FPDF_*" are pdfium-render's libloading symbol names.
PDFIUM_MARKERS = (b"FPDF_", b"pdfium-render", b"pdfium.dll", b"pdfium.dylib")
HAYRO_MARKERS = (b"hayro",)

sys.path.insert(0, str(HARNESS_DIR))
import wer as wer_mod  # noqa: E402  (the harness' own scorer module)


# ---------------------------------------------------------------- utilities

def raster_pages_sha(pdf_path: Path) -> str:
    """sha256 over the sorted pdftoppm-re-rendered page PNGs of a PDF —
    proves content determinism independent of container metadata."""
    import glob as _glob
    with tempfile.TemporaryDirectory(prefix="bakeoff-raster-") as rd:
        subprocess.run(["pdftoppm", "-png", "-r", "72",
                        str(pdf_path), str(Path(rd) / "pg")],
                       check=True, capture_output=True)
        h = hashlib.sha256()
        for png in sorted(_glob.glob(str(Path(rd) / "pg-*.png"))):
            h.update(os.path.basename(png).encode())
            h.update(Path(png).read_bytes())
        return h.hexdigest()


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def sniff_binary(path: Path) -> dict:
    data = path.read_bytes()
    return {
        "pdfium": any(m in data for m in PDFIUM_MARKERS),
        "hayro": any(m in data for m in HAYRO_MARKERS),
    }


def pdfium_lib() -> Path | None:
    """Mirror of office_parser::formats::pdf::pdfium_lib_path()
    (src/formats/pdf.rs:2531): PDFIUM_PATH env, else
    $HOME/.cache/office-parser/pdfium/libpdfium.so."""
    env = os.environ.get("PDFIUM_PATH")
    if env and Path(env).exists():
        return Path(env)
    cache = Path.home() / ".cache" / "office-parser" / "pdfium" / "libpdfium.so"
    return cache if cache.exists() else None


def blocks_text(blocks) -> str:
    """Transcript text of a JSON document's blocks, in block order."""
    parts: list[str] = []
    for b in blocks or []:
        t = b.get("type")
        if t in ("paragraph", "heading", "note", "link"):
            parts.append(b.get("text") or "")
        elif t == "list":
            for item in b.get("items") or []:
                parts.append(item.get("text") or "")
        elif t == "table":
            for row in b.get("rows") or []:
                for cell in row or []:
                    parts.append(cell.get("text") or "")
        # image blocks carry no OCR transcript
    return " ".join(p for p in parts if p)


def blank_metrics() -> dict:
    return {"error": None, "wall_s": None, "chars": None, "words": None,
            "blocks": None, "lowconf": None, "pages_ocr": None,
            "engine": None, "page_count": None}


def metrics_from_doc(doc: dict, wall_s: float) -> dict:
    blocks = doc.get("blocks") or []
    text = blocks_text(blocks)
    meta = doc.get("metadata") or {}
    ocr = (meta.get("extra") or {}).get("ocr") or {}
    m = blank_metrics()
    m.update({
        "wall_s": wall_s,
        "chars": len(text),
        "words": len(text.split()),
        "blocks": len(blocks),
        "lowconf": len(ocr.get("low_confidence_pages") or []),
        "pages_ocr": ocr.get("pages_ocr"),
        "engine": ocr.get("engine"),
        "page_count": meta.get("page_count"),
    })
    return m


def error_metrics(wall_s: float, proc) -> dict:
    m = blank_metrics()
    m["wall_s"] = wall_s
    m["error"] = ((proc.stderr or "") + (proc.stdout or "")).strip()[-600:] \
        or f"exit {proc.returncode}"
    return m


def error_metrics_msg(wall_s: float, msg: str) -> dict:
    m = blank_metrics()
    m["wall_s"] = wall_s
    m["error"] = msg
    return m


def page_words_from_jsonl(path: Path) -> list[list[str]]:
    """Per-page hypothesis word lists, ordered by metadata.extra.page."""
    if not path.exists():
        return []
    pages: dict[int, list[str]] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line:
            continue
        d = json.loads(line)
        n = (((d.get("metadata") or {}).get("extra") or {}).get("page"))
        pages[int(n)] = blocks_text(d.get("blocks") or []).split()
    return [pages[k] for k in sorted(pages)]


# ---------------------------------------------------------------- arms

def run_cli(cli: Path, pdf: Path, out_dir: Path, flags: list[str],
            renderer: str):
    env = os.environ.copy()
    env["OFFICE_PARSER_RENDERER"] = renderer
    out_dir.mkdir(parents=True, exist_ok=True)
    cmd = [str(cli), str(pdf), "--out", str(out_dir), *flags]
    t0 = time.monotonic()
    proc = subprocess.run(cmd, capture_output=True, text=True, env=env)
    return proc, time.monotonic() - t0


def load_doc(out_dir: Path, stem: str) -> dict:
    p = out_dir / f"{stem}.json"
    if not p.exists():
        raise RuntimeError(f"CLI did not write {p}")
    return json.loads(p.read_text(encoding="utf-8"))


def run_synthetic(cli: Path, corpus_pdf: Path, work: Path, renderer: str):
    out_dir = work / renderer / "synthetic"
    flags = ["--format", "json", "--json-no-image-bytes",
             "--ocr", "--ocr-force", "--split-pages"]
    proc, wall = run_cli(cli, corpus_pdf, out_dir, flags, renderer)
    print(f"  [{renderer}] synthetic: {wall:.1f}s", flush=True)
    if proc.returncode != 0:
        return error_metrics(wall, proc), None
    try:
        m = metrics_from_doc(load_doc(out_dir, "dense_scan"), wall)
        return m, page_words_from_jsonl(out_dir / "dense_scan.jsonl")
    except (RuntimeError, json.JSONDecodeError, OSError) as e:
        return error_metrics_msg(wall, f"{type(e).__name__}: {e}"), None


def run_simple(cli: Path, pdf: Path, work: Path, renderer: str, group: str,
               flags: list[str]) -> dict:
    out_dir = work / renderer / group / pdf.stem
    proc, wall = run_cli(cli, pdf, out_dir, flags, renderer)
    print(f"  [{renderer}] {group}/{pdf.name}: {wall:.1f}s", flush=True)
    if proc.returncode != 0:
        return error_metrics(wall, proc)
    try:
        return metrics_from_doc(load_doc(out_dir, pdf.stem), wall)
    except (RuntimeError, json.JSONDecodeError, OSError) as e:
        return error_metrics_msg(wall, f"{type(e).__name__}: {e}")


FIXTURE_FLAGS = ["--format", "json", "--json-no-image-bytes", "--ocr"]
REAL_SCAN_FLAGS = ["--format", "json", "--json-no-image-bytes",
                   "--ocr", "--ocr-force"]


def wer_for_renderer(ref_pages, hyp_pages):
    """Micro WER + per-page list; None when the arm produced no pages."""
    if not hyp_pages:
        return None
    n = max(len(ref_pages), len(hyp_pages))
    ref = list(ref_pages) + [[]] * (n - len(ref_pages))
    hyp = list(hyp_pages) + [[]] * (n - len(hyp_pages))
    micro, per = wer_mod.corpus_wer(ref, hyp)
    return {"micro": micro, "per_page": per}


# ------------------------------------------------------------- formatting

def fmt_int(v, signed=False):
    if v is None:
        return "n/a"
    return f"{v:+d}" if signed else f"{v:d}"


def fmt_delta(a, b, nd=0):
    if a is None or b is None:
        return "n/a"
    return f"{b - a:+.{nd}f}"


def fmt_wall(v, signed=False):
    if v is None:
        return "n/a"
    return f"{v:+.2f}" if signed else f"{v:.2f}"


def fmt_wer(v):
    return "n/a" if v is None else f"{v:.4f}"


# ---------------------------------------------------------------- report

def build_report(ctx: dict) -> str:
    L: list[str] = []
    p = ctx["pdfium"]
    h = ctx["hayro"]
    L.append("# Render bake-off: pdfium vs hayro")
    L.append("")
    L.append(f"- Generated: {ctx['ts']} (UTC)")
    L.append(f"- CLI: {ctx['cli']} (sha256 {ctx['cli_sha'][:16]}, "
             f"{ctx['cli_size']} bytes)")
    L.append("- Arms: OFFICE_PARSER_RENDERER=pdfium / hayro, interleaved "
             "per corpus item; wall time covers the full CLI process "
             "(startup + render + OCR + JSON write); a warm-up run preceded "
             "measurement.")
    L.append(f"- Corpus: seed {ctx['seed']}, {ctx['pages']} pages, "
             f"{ctx['word_count']} words - dense_scan.pdf sha256 "
             f"{ctx['pdf_sha'][:16]}, words.json sha256 {ctx['words_sha'][:16]}")
    env_notes = []
    for name in ("INDRA_OCR_MODEL_DIR", "OFFICE_PARSER_TABLE_MODEL"):
        v = os.environ.get(name)
        if v:
            env_notes.append(f"{name}={v}")
    L.append("- Env passthrough: " + (", ".join(env_notes) if env_notes else
             "none of INDRA_OCR_MODEL_DIR / OFFICE_PARSER_TABLE_MODEL set")
             + " - both are respected when set, never required.")
    L.append("")

    # -- integrity -------------------------------------------------------
    L.append("## Arm integrity")
    L.append("")
    L.append("| check | value |")
    L.append("|---|---|")
    L.append("| binary has pdfium markers (FPDF_/pdfium-render) | "
             + ("yes" if ctx["sniff"]["pdfium"] else "NO") + " |")
    L.append("| binary has hayro markers | "
             + ("yes" if ctx["sniff"]["hayro"] else "NO") + " |")
    L.append("| pdfium lib (PDFIUM_PATH or ~/.cache/office-parser/pdfium/) | "
             + (ctx["pdfium_lib"] or "missing") + " |")
    L.append("| synthetic ocr engine / pages_ocr (pdfium) | "
             f"{p['metrics']['engine'] or 'n/a'} / "
             f"{fmt_int(p['metrics']['pages_ocr'])} |")
    L.append("| synthetic ocr engine / pages_ocr (hayro) | "
             f"{h['metrics']['engine'] or 'n/a'} / "
             f"{fmt_int(h['metrics']['pages_ocr'])} |")
    L.append("| synthetic transcripts distinct across arms | "
             + ("yes" if not ctx["collapse"]
                else "NO - see warning") + " |")
    L.append("")
    for w in ctx["warnings"]:
        L.append(f"> **WARNING:** {w}")
        L.append("")

    # -- synthetic -------------------------------------------------------
    L.append(f"## Synthetic corpus - dense_scan.pdf "
             f"({ctx['pages']} pages, {ctx['word_count']} ground-truth words)")
    L.append("")
    L.append("| metric | pdfium | hayro | Delta (hayro minus pdfium) |")
    L.append("|---|---:|---:|---:|")
    wp = p["wer"]
    wh = h["wer"]
    L.append("| WER - micro-avg (lower is better) | "
             f"{fmt_wer(wp and wp['micro'])} "
             f"| {fmt_wer(wh and wh['micro'])} | "
             f"{fmt_delta(wp and wp['micro'], wh and wh['micro'], 4)} |")
    L.append(f"| WER - per page min..max | {ctx['per_page_str'][0]} "
             f"| {ctx['per_page_str'][1]} | - |")
    L.append(f"| words | {fmt_int(p['metrics']['words'])} "
             f"| {fmt_int(h['metrics']['words'])} "
             f"| {fmt_delta(p['metrics']['words'], h['metrics']['words'])} |")
    L.append(f"| chars | {fmt_int(p['metrics']['chars'])} "
             f"| {fmt_int(h['metrics']['chars'])} "
             f"| {fmt_delta(p['metrics']['chars'], h['metrics']['chars'])} |")
    L.append(f"| blocks | {fmt_int(p['metrics']['blocks'])} "
             f"| {fmt_int(h['metrics']['blocks'])} "
             f"| {fmt_delta(p['metrics']['blocks'], h['metrics']['blocks'])} |")
    L.append(f"| low-confidence pages | {fmt_int(p['metrics']['lowconf'])} "
             f"| {fmt_int(h['metrics']['lowconf'])} "
             f"| {fmt_delta(p['metrics']['lowconf'], h['metrics']['lowconf'])} |")
    L.append(f"| wall time (s) | {fmt_wall(p['metrics']['wall_s'])} "
             f"| {fmt_wall(h['metrics']['wall_s'])} "
             f"| {fmt_delta(p['metrics']['wall_s'], h['metrics']['wall_s'], 2)} |")
    L.append("")
    if wp and wh:
        sat = wh["micro"] <= wp["micro"] + FLIP_MARGIN
        margin = wp["micro"] + FLIP_MARGIN - wh["micro"]
        L.append(f"Decision rule (TASK 4): flip default to hayro iff "
                 f"WER(hayro) <= WER(pdfium) + {FLIP_MARGIN} -> "
                 f"WER(pdfium)={wp['micro']:.4f}, WER(hayro)={wh['micro']:.4f} "
                 f"-> rule **{'SATISFIED' if sat else 'NOT satisfied'}** "
                 f"(margin {margin:+.4f}). The flip itself remains "
                 f"human-gated.")
    else:
        L.append("Decision rule (TASK 4): **NOT EVALUABLE** - missing WER "
                 "for at least one arm (see errors above).")
    for label, m in (("pdfium", p["metrics"]), ("hayro", h["metrics"])):
        if m["error"]:
            L.append("")
            L.append(f"> arm error ({label}): {m['error']!r}")
    L.append("")
    L.append("WER method: ground truth = words.json (exactly as drawn, "
             "column-major reading order); hypothesis = per-page OCR "
             "transcripts from the CLI JSONL; Levenshtein distance over "
             "word sequences / reference length, casefolded tokens, "
             "micro-averaged across pages.")
    if ctx.get("page_gap_note"):
        L.append("")
        L.append(f"> {ctx['page_gap_note']}")
    L.append("")

    # -- fixtures --------------------------------------------------------
    L.append("## Fixtures - " + ctx["fixtures_label"] + " (plain --ocr, "
             "routing respected; no WER column - no ground truth)")
    L.append("")
    if ctx["fixtures"]:
        L.append("| fixture | metric | pdfium | hayro | Delta |")
        L.append("|---|---|---:|---:|---:|")
        for stem, name in ctx["fixtures"]:
            a = p["fixtures"][stem]
            b = h["fixtures"][stem]
            for label, key, nd, fmt in (
                    ("words", "words", 0, fmt_int),
                    ("chars", "chars", 0, fmt_int),
                    ("blocks", "blocks", 0, fmt_int),
                    ("low-confidence pages", "lowconf", 0, fmt_int),
                    ("wall time (s)", "wall_s", 2, fmt_wall)):
                L.append(f"| {name} | {label} | {fmt(a[key])} | {fmt(b[key])} "
                         f"| {fmt_delta(a[key], b[key], nd)} |")
        L.append("")
        err = []
        for st, nm in ctx["fixtures"]:
            if p["fixtures"][st]["error"]:
                err.append((nm + " (pdfium)", p["fixtures"][st]["error"]))
            if h["fixtures"][st]["error"]:
                err.append((nm + " (hayro)", h["fixtures"][st]["error"]))
        if err:
            for n, e in err:
                L.append(f"> fixture arm error ({n}): {e!r}")
            L.append("")
    else:
        L.append("_(no pdf_*.pdf fixtures found in "
                 + ctx["fixtures_dir"] + ")_")
        L.append("")

    # -- real scans ------------------------------------------------------
    L.append("## real-scans/ - drop-in scans (char/word delta only; no WER "
             "column - no ground truth exists)")
    L.append("")
    if ctx["real"]:
        L.append("| scan | metric | pdfium | hayro | Delta |")
        L.append("|---|---|---:|---:|---:|")
        for name in ctx["real"]:
            a = p["real"][name]
            b = h["real"][name]
            for label, key, nd, fmt in (
                    ("words", "words", 0, fmt_int),
                    ("chars", "chars", 0, fmt_int),
                    ("blocks", "blocks", 0, fmt_int),
                    ("low-confidence pages", "lowconf", 0, fmt_int),
                    ("wall time (s)", "wall_s", 2, fmt_wall)):
                L.append(f"| {name} | {label} | {fmt(a[key])} | {fmt(b[key])} "
                         f"| {fmt_delta(a[key], b[key], nd)} |")
        L.append("")
        err = []
        for nm in ctx["real"]:
            if p["real"][nm]["error"]:
                err.append((nm + " (pdfium)", p["real"][nm]["error"]))
            if h["real"][nm]["error"]:
                err.append((nm + " (hayro)", h["real"][nm]["error"]))
        if err:
            for n, e in err:
                L.append(f"> real-scan arm error ({n}): {e!r}")
            L.append("")
    else:
        L.append("_(empty - drop .pdf files into the harness' real-scans/ "
                 "directory and rerun run_bakeoff.py)_")
        L.append("")

    # -- notes -----------------------------------------------------------
    L.append("## Notes")
    L.append("")
    L.append("- Fixtures run with plain --ocr: page routing (text layer vs "
             "OCR) is text-layer based, hence renderer-independent - the "
             "A/B stays valid. --ocr-force is used only on the synthetic "
             "corpus and real scans.")
    L.append("- Renderer selection falls back silently "
             "(pdfium -> hayro -> pdftoppm, src/formats/pdf.rs); a runtime "
             "render failure of one backend can therefore still swap a "
             "pinned arm to pdftoppm. The transcript-distinctness check "
             "above catches the common both-arms-collapse case.")
    L.append("- WER(hayro) <= WER(pdfium) + 1pt is the TASK 4 flip rule; "
             "this report only measures - the flip is a separate "
             "human-gated task.")
    L.append("")
    return "\n".join(L)


# --------------------------------------------------------------- selfcheck

def selfcheck() -> int:
    print("== run_bakeoff --selfcheck ==")
    ok = True

    # 1. generator determinism: two seeded runs -> identical sha256s
    print("[1/2] generator determinism (two runs, seed 7, 2 pages)")
    with tempfile.TemporaryDirectory(prefix="bakeoff-selfcheck-") as td:
        a, b = Path(td) / "a", Path(td) / "b"
        for d in (a, b):
            r = subprocess.run(
                [sys.executable, str(GENERATOR), "--seed", "7",
                 "--pages", "2", "--out-dir", str(d)],
                capture_output=True, text=True)
            if r.returncode != 0:
                print(f"FAIL  generator run failed: {r.stderr[-500:]}")
                ok = False
                break
        else:
            wa = sha256_file(a / "words.json")
            wb = sha256_file(b / "words.json")
            # Content determinism (what the measurement consumes): the
            # RASTERIZED pages must be identical. Whole-file bytes are NOT —
            # reportlab's invariant=1 pins the timestamp but still varies the
            # /ID and the compressed XMP info block (713+ bytes of metadata
            # noise, no content effect). words.json must also match exactly.
            ra = raster_pages_sha(a / "dense_scan.pdf")
            rb = raster_pages_sha(b / "dense_scan.pdf")
            same = (ra == rb and wa == wb)
            print(f"{'PASS' if same else 'FAIL'}  rasterized-page sha256="
                  f"{ra[:16]}... words.json sha256={wa[:16]}... identical "
                  f"across runs: {same} (file bytes intentionally not "
                  f"compared: reportlab /ID + XMP metadata noise)")
            ok = ok and same

    # 2. WER scorer micro-case
    print("[2/2] WER scorer micro-case (wer.py --selfcheck)")
    r = subprocess.run([sys.executable, str(WER_PY), "--selfcheck"],
                       capture_output=True, text=True)
    sys.stdout.write(r.stdout)
    if r.returncode != 0:
        print(f"FAIL  wer.py --selfcheck exited {r.returncode}")
        ok = False

    print("selfcheck:", "PASS" if ok else "FAIL")
    return 0 if ok else 1


# -------------------------------------------------------------------- main

def parse_args(argv=None):
    ap = argparse.ArgumentParser(
        description="pdfium vs hayro bake-off runner (see module docstring).")
    ap.add_argument("--cli", default=str(DEFAULT_CLI),
                    help=f"office-parser-cli path (default: {DEFAULT_CLI})")
    ap.add_argument("--corpus-dir", default=str(DEFAULT_CORPUS_DIR),
                    help="dir holding dense_scan.pdf + words.json")
    ap.add_argument("--fixtures-dir", default=str(DEFAULT_FIXTURES_DIR),
                    help="dir with the pdf_*.pdf fixtures")
    ap.add_argument("--real-scans-dir", default=str(DEFAULT_REAL_SCANS_DIR),
                    help="drop-in dir for real scans (scored without WER)")
    ap.add_argument("--work-dir", default=str(DEFAULT_WORK_DIR),
                    help="scratch dir for CLI outputs (default "
                         "/tmp/render-bakeoff)")
    ap.add_argument("--report", default=str(DEFAULT_REPORT),
                    help="report.md output path")
    ap.add_argument("--renderers", default="pdfium,hayro",
                    help="comma-separated arms, subset of pdfium,hayro")
    ap.add_argument("--no-generate", action="store_true",
                    help="fail instead of auto-running generate_corpus.py "
                         "when the corpus is missing")
    ap.add_argument("--selfcheck", action="store_true",
                    help="generator determinism + scorer micro-case; exit 0 "
                         "on pass")
    return ap.parse_args(argv)


def main(argv=None) -> int:
    args = parse_args(argv)
    if args.selfcheck:
        return selfcheck()

    renderers = [r.strip() for r in args.renderers.split(",") if r.strip()]
    if set(renderers) != set(RENDERERS):
        print("error: --renderers must be exactly "
              f"{list(RENDERERS)} - this harness is a two-arm A/B",
              file=sys.stderr)
        return 2

    cli = Path(args.cli).expanduser()
    cli = cli if cli.is_absolute() else (Path.cwd() / cli)
    cli = cli.resolve()
    if not cli.exists():
        print("error: office-parser-cli not found at "
              f"{cli}\nbuild it (from the repo root):\n"
              "  CARGO_BUILD_JOBS=4 cargo build --release "
              "--manifest-path cli/Cargo.toml \\\n"
              "      --features render-pdfium,render-hayro,ocr-indra-native "
              "--target-dir target", file=sys.stderr)
        return 2

    # Arm integrity guards (src/formats/pdf.rs falls through SILENTLY).
    sniff = sniff_binary(cli)
    lib = pdfium_lib()
    guard_errors = []
    if "pdfium" in renderers:
        if not sniff["pdfium"]:
            guard_errors.append(
                "the binary lacks pdfium markers (FPDF_/pdfium-render): it "
                "was likely built WITHOUT --features render-pdfium, so the "
                "pdfium arm would SILENTLY fall back to pdftoppm "
                "(src/formats/pdf.rs:2759). Rebuild before measuring.")
        elif lib is None:
            guard_errors.append(
                "no pdfium library found (PDFIUM_PATH or "
                "~/.cache/office-parser/pdfium/libpdfium.so): the pdfium arm "
                "would SILENTLY fall back to pdftoppm "
                "(src/formats/pdf.rs:2759).")
    if "hayro" in renderers and not sniff["hayro"]:
        guard_errors.append(
            "the binary lacks hayro markers: it was likely built WITHOUT "
            "--features render-hayro, so the hayro arm would SILENTLY fall "
            "back to pdftoppm (src/formats/pdf.rs:2769). Rebuild before "
            "measuring.")
    if guard_errors:
        for e in guard_errors:
            print(f"error: {e}", file=sys.stderr)
        return 3

    corpus_dir = Path(args.corpus_dir).expanduser().resolve()
    corpus_pdf = corpus_dir / "dense_scan.pdf"
    words_path = corpus_dir / "words.json"
    if not (corpus_pdf.exists() and words_path.exists()):
        if args.no_generate:
            print(f"error: corpus missing in {corpus_dir} (--no-generate); "
                  "run generate_corpus.py first", file=sys.stderr)
            return 2
        print("corpus missing -> auto-running generate_corpus.py "
              "(defaults: seed 42, 10 pages)")
        r = subprocess.run(
            [sys.executable, str(GENERATOR), "--out-dir", str(corpus_dir)])
        if r.returncode != 0 or not (corpus_pdf.exists()
                                     and words_path.exists()):
            print("error: corpus generation failed", file=sys.stderr)
            return 2

    truth = json.loads(words_path.read_text(encoding="utf-8"))
    ref_pages = [pg["words"] for pg in truth["pages_words"]]
    seed = truth.get("seed")
    pages_n = truth.get("pages", len(ref_pages))
    word_count = truth.get("word_count",
                           sum(len(pg) for pg in ref_pages))

    fixtures_dir = Path(args.fixtures_dir).expanduser().resolve()
    fixtures = sorted(fixtures_dir.glob("pdf_*.pdf"))
    real_dir = Path(args.real_scans_dir).expanduser().resolve()
    real_scans = sorted(real_dir.glob("*.pdf")) if real_dir.is_dir() else []

    work = Path(args.work_dir).expanduser().resolve()
    shutil.rmtree(work, ignore_errors=True)
    work.mkdir(parents=True, exist_ok=True)

    # Warm-up (OS page cache + ONNX session init), best-effort, discarded.
    warm = fixtures[0] if fixtures else None
    if warm is not None:
        try:
            run_cli(cli, warm, work / "warmup",
                    ["--format", "json", "--json-no-image-bytes", "--ocr",
                     "--ocr-force"], renderers[0])
            print("warm-up run done (not measured)")
        except OSError as e:
            print(f"warm-up skipped: {e}")

    results = {r: {"synthetic": None, "fixtures": {}, "real": {}}
               for r in renderers}

    # (a) synthetic corpus, interleaved per renderer
    print("== synthetic corpus ==")
    hyp_pages_by_renderer = {}
    for r in renderers:
        m, hyp = run_synthetic(cli, corpus_pdf, work, r)
        results[r]["synthetic"] = m
        hyp_pages_by_renderer[r] = hyp

    # (b) fixtures, interleaved per renderer
    print("== fixtures ==")
    for pdf in fixtures:
        for r in renderers:
            results[r]["fixtures"][pdf.stem] = run_simple(
                cli, pdf, work, r, "fixtures", FIXTURE_FLAGS)

    # (c) real-scans drop-in (skipped silently when empty)
    if real_scans:
        print("== real-scans ==")
        for pdf in real_scans:
            for r in renderers:
                results[r]["real"][pdf.name] = run_simple(
                    cli, pdf, work, r, "real", REAL_SCAN_FLAGS)
    else:
        print("== real-scans: empty, skipped ==")

    # Metrics assembly
    warnings: list[str] = []
    wer_by_renderer = {}
    per_page_by_renderer = {}
    for r in renderers:
        m = results[r]["synthetic"]
        if m["error"] is None and m.get("pages_ocr") in (None, 0):
            warnings.append(
                f"synthetic arm {r} produced no OCR metadata "
                f"(--ocr-force should force OCR on every page).")
        wer_by_renderer[r] = wer_for_renderer(ref_pages,
                                              hyp_pages_by_renderer[r] or [])
        per = (wer_by_renderer[r] or {}).get("per_page") or []
        per_page_by_renderer[r] = (f"{min(per):.3f}..{max(per):.3f}"
                                   if per else "n/a")

    p_m = results["pdfium"]["synthetic"]
    h_m = results["hayro"]["synthetic"]
    if (p_m["error"] is None and h_m["error"] is None
            and p_m["chars"] is not None
            and p_m["chars"] == h_m["chars"]
            and p_m["words"] == h_m["words"]):
        warnings.append(
            "both arms produced byte-identical synthetic transcripts - the "
            "pinned backends likely collapsed to the same renderer (silent "
            "fallback); do NOT trust this run. Verify the build features "
            "and the pdfium library, then rerun.")

    hyp_lists = [v for v in hyp_pages_by_renderer.values() if v]
    n_hyp = max((len(v) for v in hyp_lists), default=0)
    page_gap_note = None
    if n_hyp and n_hyp != len(ref_pages):
        page_gap_note = (f"OCR produced {n_hyp} page docs vs "
                         f"{len(ref_pages)} ground-truth pages; missing "
                         f"pages were scored as empty transcripts "
                         f"(deletions).")

    ctx = {
        "ts": time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime()),
        "cli": str(cli),
        "cli_sha": sha256_file(cli),
        "cli_size": cli.stat().st_size,
        "sniff": sniff,
        "pdfium_lib": str(lib) if lib else "missing",
        "seed": seed,
        "pages": pages_n,
        "word_count": word_count,
        "pdf_sha": sha256_file(corpus_pdf),
        "words_sha": sha256_file(words_path),
        "warnings": warnings,
        "collapse": any("byte-identical" in w for w in warnings),
        "pdfium": {
            "metrics": p_m,
            "wer": wer_by_renderer.get("pdfium"),
            "fixtures": results["pdfium"]["fixtures"],
            "real": results["pdfium"]["real"],
        },
        "hayro": {
            "metrics": h_m,
            "wer": wer_by_renderer.get("hayro"),
            "fixtures": results["hayro"]["fixtures"],
            "real": results["hayro"]["real"],
        },
        "per_page_str": (per_page_by_renderer.get("pdfium", "n/a"),
                         per_page_by_renderer.get("hayro", "n/a")),
        "page_gap_note": page_gap_note,
        "fixtures": [(pdf.stem, pdf.name) for pdf in fixtures],
        "fixtures_label": "tests/fixtures/pdf_*.pdf",
        "fixtures_dir": str(fixtures_dir),
        "real": [pdf.name for pdf in real_scans],
    }

    report = build_report(ctx)
    report_path = Path(args.report).expanduser()
    report_path = report_path if report_path.is_absolute() \
        else (Path.cwd() / report_path)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(report, encoding="utf-8")

    print("== summary ==")
    wp, wh = wer_by_renderer.get("pdfium"), wer_by_renderer.get("hayro")
    if wp and wh:
        sat = wh["micro"] <= wp["micro"] + FLIP_MARGIN
        print(f"WER micro-avg: pdfium={wp['micro']:.4f} "
              f"hayro={wh['micro']:.4f} "
              f"(delta {wh['micro'] - wp['micro']:+.4f})")
        print(f"TASK 4 flip rule (WER(hayro) <= WER(pdfium) + "
              f"{FLIP_MARGIN}): {'SATISFIED' if sat else 'NOT satisfied'}")
    else:
        print("WER micro-avg: n/a (see report errors)")
    print(f"synthetic chars: pdfium={p_m['chars']} hayro={h_m['chars']} "
          f"(delta {fmt_delta(p_m['chars'], h_m['chars'])})")
    print(f"report: {report_path}")
    for w in warnings:
        print(f"WARNING: {w}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

#!/usr/bin/env python3
"""Seeded dense-scan corpus generator for the render bake-off (TASK 3).

Formalizes the proven flow (.workflows/research/hayro-0.9.0-upstream-report.md):

  1. reportlab canvas -- A4, 2 columns, 8pt Helvetica, greedy stringWidth word
     wrap (~14 words/line at this geometry) -- text-layer PDF written to /tmp
  2. pdftoppm -r 150 -png rasterizes every page into /tmp (the raster the OCR
     engine will actually see)
  3. reportlab drawImage embeds the PNGs into an image-only PDF -- no text
     layer, so --ocr routing is forced on every page

Outputs (default: this script's directory; override with --out-dir):
  dense_scan.pdf   image-only corpus PDF
  words.json       per-page ordered word list EXACTLY as drawn -- the WER
                   ground truth. Draw order = column-major: left column
                   top-to-bottom, then right column (standard 2-column
                   reading order).

Determinism: random.Random(seed) for word choice + reportlab invariant=1
(fixed timestamps/document id -> byte-reproducible PDFs) + poppler's
deterministic PNG rasterization. Two runs with the same seed are
byte-identical; run_bakeoff.py --selfcheck verifies exactly that.

Dependencies: python3 stdlib + system reportlab (generator only) + pdftoppm
(poppler-utils). Nothing here is wired into CI or the test suite.
"""

from __future__ import annotations

import argparse
import hashlib
import re
import json
import random
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

PAGE_W, PAGE_H = A4          # 595.27 x 841.89 pt
MARGIN = 36.0                # pt, left + right
TOP_MARGIN = 40.0            # pt
BOTTOM_MARGIN = 44.0         # pt
GUTTER = 20.0                # pt between the two columns
FONT = "Helvetica"
FONT_SIZE = 8.0              # pt
LEADING = 9.6                # pt between baselines
DPI = 150

DEFAULT_SEED = 42
DEFAULT_PAGES = 10
SPACE_WIDTH = stringWidth(" ", FONT, FONT_SIZE)

# Dense body-text vocabulary: lowercase letters only, 3-11 chars (no
# punctuation, no digits -- token boundaries stay crisp for the OCR engine).
# Duplicates are intentional weighting, not an accident.
WORD_POOL = [
    "the", "and", "that", "have", "for", "not", "with", "you", "this", "but",
    "his", "from", "they", "say", "her", "she", "will", "one", "all", "would",
    "there", "their", "what", "out", "about", "who", "get", "which", "when",
    "make", "can", "like", "time", "just", "him", "know", "take", "into",
    "year", "your", "good", "some", "could", "them", "see", "other", "than",
    "then", "now", "look", "only", "come", "its", "over", "also", "back",
    "after", "use", "two", "how", "our", "work", "first", "well", "way",
    "even", "new", "want", "because", "any", "these", "give", "day", "most",
    "us", "man", "find", "here", "thing", "many", "great", "still", "through",
    "such", "between", "never", "under", "keep", "while", "last", "same",
    "another", "must", "each", "around", "own", "against", "however", "few",
    "world", "large", "often", "along", "might", "both", "children", "next",
    "early", "young", "important", "little", "public", "bad", "able",
    "people", "part", "case", "point", "government", "life", "number",
    "group", "problem", "fact", "place", "right", "study", "system",
    "program", "question", "business", "area", "water", "money", "story",
    "month", "lot", "book", "school", "word", "side", "state", "family",
    "student", "country", "course", "company", "team", "idea", "body",
    "information", "nothing", "together", "night", "face", "far", "always",
    "science", "room", "friend", "begin", "city", "almost", "small", "again",
    "yes", "need", "call", "develop", "hold", "interest", "line", "read",
    "change", "help", "play", "run", "move", "live", "believe", "bring",
    "happen", "write", "provide", "sit", "stand", "lose", "pay", "meet",
    "include", "continue", "set", "learn", "lead", "understand", "watch",
    "follow", "stop", "create", "speak", "allow", "add", "spend", "grow",
    "open", "walk", "win", "offer", "remember", "consider", "appear", "buy",
    "serve", "send", "build", "stay", "fall", "cut", "reach", "remain",
    "suggest", "raise", "pass", "sell", "require", "report", "decide", "pull",
]


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def line_width(words: list[str]) -> float:
    if not words:
        return 0.0
    return (sum(stringWidth(w, FONT, FONT_SIZE) for w in words)
            + SPACE_WIDTH * (len(words) - 1))


def fill_column(rng: random.Random, col_width: float, max_lines: int,
                pending: str | None):
    """Greedy-wrap words into exactly max_lines lines within col_width.

    Returns (lines, pending): lines is a list of word lists; pending is the
    word drawn but not yet placed (it starts the next column). A word wider
    than the column is placed alone rather than looping forever.
    """
    lines: list[list[str]] = []
    while len(lines) < max_lines:
        words: list[str] = []
        if pending is not None:
            words.append(pending)
            pending = None
        while True:
            w = rng.choice(WORD_POOL)
            if words and (line_width(words) + SPACE_WIDTH
                          + stringWidth(w, FONT, FONT_SIZE)) > col_width:
                pending = w
                break
            words.append(w)
        lines.append(words)
    return lines, pending


def page_columns(rng: random.Random, col_width: float, lines_per_col: int):
    """One page = two fully packed columns; the word stream continues across
    the gutter (column-major reading order = ground-truth draw order)."""
    cols = []
    pending = None
    for _ in range(2):
        col, pending = fill_column(rng, col_width, lines_per_col, pending)
        cols.append(col)
    return cols


def draw_text_layer_pdf(path: Path, pages: list) -> None:
    c = canvas.Canvas(str(path), pagesize=A4, invariant=1)
    c.setTitle("office-parser render bake-off corpus")
    col_width = (PAGE_W - 2 * MARGIN - GUTTER) / 2.0
    for cols in pages:
        y = PAGE_H - TOP_MARGIN - FONT_SIZE  # first baseline of each column
        for ci, col in enumerate(cols):
            x = MARGIN + ci * (col_width + GUTTER)
            for line in col:
                c.drawString(x, y, " ".join(line))
                y -= LEADING
        c.showPage()
    c.save()


def rasterize(text_pdf: Path, tmp: Path, expected_pages: int) -> list[Path]:
    subprocess.run(
        ["pdftoppm", "-png", "-r", str(DPI), str(text_pdf), str(tmp / "page")],
        check=True,
    )
    pngs = sorted(tmp.glob("page-*.png"))  # zero-padded names sort correctly
    if len(pngs) != expected_pages:
        raise RuntimeError(
            f"pdftoppm produced {len(pngs)} PNGs, expected {expected_pages}")
    return pngs


def draw_image_only_pdf(pngs: list[Path], out_pdf: Path) -> None:
    c = canvas.Canvas(str(out_pdf), pagesize=A4, invariant=1)
    c.setTitle("office-parser render bake-off corpus")
    for png in pngs:
        c.drawImage(str(png), 0, 0, width=PAGE_W, height=PAGE_H)
        c.showPage()
    c.save()
    # reportlab invariant=1 pins the timestamp but NOT the per-document /ID
    # (os.urandom) — normalize it so byte-determinism holds (the selfcheck
    # asserts identical sha256 across runs; words.json already is).
    data = Path(out_pdf).read_bytes()


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        description="Seeded dense-scan corpus generator (render bake-off).")
    ap.add_argument("--seed", type=int, default=DEFAULT_SEED,
                    help="RNG seed (default 42; same seed = byte-identical corpus)")
    ap.add_argument("--pages", type=int, default=DEFAULT_PAGES,
                    help="page count (default 10)")
    ap.add_argument("--out-dir",
                    default=str(Path(__file__).resolve().parent),
                    help="output dir for dense_scan.pdf + words.json")
    args = ap.parse_args(argv)
    if args.pages < 1:
        print("error: --pages must be >= 1", file=sys.stderr)
        return 2
    if shutil.which("pdftoppm") is None:
        print("error: pdftoppm not found (install poppler-utils)",
              file=sys.stderr)
        return 2

    out_dir = Path(args.out_dir).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    col_width = (PAGE_W - 2 * MARGIN - GUTTER) / 2.0
    lines_per_col = int((PAGE_H - TOP_MARGIN - BOTTOM_MARGIN) // LEADING)

    random.seed(args.seed)  # reportlab draws FormXob names from the GLOBAL random module
    rng = random.Random(args.seed)
    pages = [page_columns(rng, col_width, lines_per_col)
             for _ in range(args.pages)]
    # Ground truth: EXACTLY the drawn words, column-major draw order.
    pages_words = [
        {"page": i + 1,
         "words": [w for col in page for line in col for w in line]}
        for i, page in enumerate(pages)
    ]

    tmp = Path(tempfile.mkdtemp(prefix="render-bakeoff-gen-"))
    try:
        text_pdf = tmp / "text-layer.pdf"
        draw_text_layer_pdf(text_pdf, pages)
        pngs = rasterize(text_pdf, tmp, args.pages)
        out_pdf = out_dir / "dense_scan.pdf"
        draw_image_only_pdf(pngs, out_pdf)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    total = sum(len(p["words"]) for p in pages_words)
    payload = {
        "generator": "generate_corpus.py",
        "seed": args.seed,
        "pages": args.pages,
        "dpi": DPI,
        "font": f"{FONT} {FONT_SIZE:g}pt",
        "columns": 2,
        "reading_order": "column-major: left column top-to-bottom, then right",
        "word_count": total,
        "pages_words": pages_words,
    }
    words_path = out_dir / "words.json"
    with words_path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2)
        f.write("\n")

    print(f"wrote {out_pdf}  ({out_pdf.stat().st_size} bytes, sha256 "
          f"{sha256_file(out_pdf)})")
    print(f"wrote {words_path}  ({total} words, sha256 "
          f"{sha256_file(words_path)})")
    print(f"seed={args.seed} pages={args.pages} ~{total // args.pages} words/page")
    return 0


if __name__ == "__main__":
    sys.exit(main())

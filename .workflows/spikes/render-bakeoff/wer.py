#!/usr/bin/env python3
"""Word Error Rate (WER) scorer for the render bake-off (TASK 3).

WER = standard Levenshtein edit distance over the word sequence, normalized
by the reference length. Tokens are whitespace-separated; both sides are
casefolded by default (OCR case drift is renderer-neutral noise) --
--no-casefold gives exact matching.

Reference: words.json from generate_corpus.py (exactly as drawn).
Hypothesis: the OCR transcript extracted from the CLI's JSON output.

python3 stdlib only.

Self-check headline (hand-computed):
  ref ["a","b","c"] vs hyp ["a","x","c"]  ->  WER = 1/3 = 0.333333
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


def edit_distance(a: list[str], b: list[str]) -> int:
    """Levenshtein distance: insert/delete/substitute, each cost 1."""
    if not a:
        return len(b)
    if not b:
        return len(a)
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(
                prev[j] + 1,                 # deletion
                cur[j - 1] + 1,              # insertion
                prev[j - 1] + (ca != cb),    # substitution
            ))
        prev = cur
    return prev[-1]


def _norm(words, casefold: bool) -> list[str]:
    return [w.casefold() for w in words] if casefold else list(words)


def wer(ref: list[str], hyp: list[str], casefold: bool = True) -> float:
    """WER for one reference/hypothesis word sequence.

    An empty reference scores 0.0 for an empty hypothesis, else 1.0.
    """
    r, h = _norm(ref, casefold), _norm(hyp, casefold)
    if not r:
        return 0.0 if not h else 1.0
    return edit_distance(r, h) / len(r)


def corpus_wer(ref_pages, hyp_pages, casefold: bool = True):
    """Micro-averaged WER over aligned page word lists.

    ref_pages / hyp_pages: equal-length lists of per-page word lists (pad a
    missing hypothesis page with [] before calling). Returns
    (micro_wer, per_page_wer_list). Micro-average = total edit distance /
    total reference words -- page-size independent, standard for corpora.
    """
    per = [wer(r, h, casefold) for r, h in zip(ref_pages, hyp_pages)]
    errors = sum(edit_distance(_norm(r, casefold), _norm(h, casefold))
                 for r, h in zip(ref_pages, hyp_pages))
    total = sum(len(_norm(r, casefold)) for r in ref_pages)
    micro = errors / total if total else 0.0
    return micro, per


# (ref, hyp, expected, why) -- every expected value hand-computed.
CASES = [
    (["a", "b", "c"], ["a", "x", "c"], 1 / 3, "1 substitution / 3 words"),
    (["a", "b", "c"], ["a", "b", "c"], 0.0, "exact match"),
    (["a", "b"], ["b", "a"], 1.0, "swap = 2 substitutions / 2 words"),
    (["a", "b", "c", "d"], ["a", "b"], 0.5, "2 deletions / 4 words"),
    (["a", "b"], ["a", "b", "c"], 0.5, "1 insertion / 2 words"),
    (["the", "cat", "sat"], ["the", "cat", "sat", "on", "mat"], 2 / 3,
     "2 insertions / 3 words"),
    ([], [], 0.0, "both empty"),
    ([], ["a"], 1.0, "empty reference, non-empty hypothesis"),
    (["The", "Cat"], ["the", "cat"], 0.0, "casefold is the default"),
]


def selfcheck() -> int:
    fails = 0
    for ref, hyp, expected, why in CASES:
        got = wer(ref, hyp)
        ok = abs(got - expected) <= 1e-9
        fails += not ok
        print(f"{'PASS' if ok else 'FAIL'}  wer({ref!r}, {hyp!r}) = "
              f"{got:.6f}  (expected {expected:.6f}; {why})")
    print()
    print("WER self-check headline: ref=['a','b','c'] vs hyp=['a','x','c'] "
          "-> WER = 1/3 = 0.333333")
    print("selfcheck:", "PASS" if not fails else f"FAIL ({fails} case(s))")
    return 0 if not fails else 1


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        description="WER scorer for the render bake-off.")
    ap.add_argument("--ref", help="ground-truth words.json (generate_corpus.py)")
    ap.add_argument("--page", type=int, default=None,
                    help="score only this 1-based page (default: all pages)")
    ap.add_argument("--hyp", help="transcript file (whitespace-split words)")
    ap.add_argument("--no-casefold", action="store_true",
                    help="exact-case token matching")
    ap.add_argument("--selfcheck", action="store_true",
                    help="assert hand-computed WER values and exit")
    args = ap.parse_args(argv)
    if args.selfcheck:
        return selfcheck()
    if not (args.ref and args.hyp):
        ap.error("--ref and --hyp are required (or use --selfcheck)")
    doc = json.loads(Path(args.ref).read_text(encoding="utf-8"))
    pages = doc["pages_words"]
    if args.page is not None:
        ref_words = next(p["words"] for p in pages if p["page"] == args.page)
    else:
        ref_words = [w for p in pages for w in p["words"]]
    hyp_words = Path(args.hyp).read_text(encoding="utf-8").split()
    value = wer(ref_words, hyp_words, casefold=not args.no_casefold)
    print(f"WER = {value:.6f}  ({len(hyp_words)} hyp words vs "
          f"{len(ref_words)} ref words)")
    return 0


if __name__ == "__main__":
    sys.exit(main())

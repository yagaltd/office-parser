# CLI Usage

The canonical CLI reference lives in [cli/README.md](../cli/README.md).

Use that page for:

- installation
- examples
- flags
- output layout
- JSON export options

## MorphEditor markdown output

`--format markdown` emits MorphEditor-compatible markdown wrapped in OKF
frontmatter. Presentation extras:

- `slides: <ratio>` is derived from the slide size (16:9 / 4:3).
- `--slide-transition <up|down|left|right|fade>` sets the
  `transition:` frontmatter key (validated against MorphEditor's
  `SLIDE_TRANSITIONS`).
- Speaker notes, `::: columns`, `::: table {widths=[..]}` and image
  media lines are emitted where the source carries them.

Note: `--chunk-size` output is the CognitiveOS chunked view and is NOT
MorphEditor markdown.

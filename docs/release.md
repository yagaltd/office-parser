# Release Guide

## Pre-release checklist

1. Update versions in `office-parser/Cargo.toml` and `office-parser/cli/Cargo.toml`.
2. Ensure README/docs are aligned with implementation.
3. Run tests and packaging checks.

## Verification commands

```bash
cargo test -p office-parser --tests
cargo test -p office-parser-cli --tests
cargo check -p office-parser --no-default-features
cargo check -p office-parser --features "json,xml"
cargo package -p office-parser --allow-dirty --no-verify
cargo package -p office-parser-cli --allow-dirty --no-verify
```

## Publish order

1. Publish `office-parser`.
2. Wait for availability on crates.io.
3. Publish `office-parser-cli` with matching `office-parser` version dependency.

## Post-publish checks

```bash
cargo install office-parser-cli
office-parser-cli --help
```

Confirm docs.rs builds for both crates.

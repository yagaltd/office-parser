//! Mid-parse progress lines.
//!
//! The library never writes to a process stream on its own: callers that want
//! the "rendering N page(s) at 200 dpi" / "ocr page 3/40" chatter install a
//! sink (the CLI installs one that prints to stderr). With no sink installed
//! parsing is silent, which is what embedders — `CognitiveOS` V3, `MorphEditor` —
//! need.

use std::sync::Mutex;

/// A line-oriented progress sink. A plain `fn` pointer keeps [`set_sink`]
/// callable from anywhere including tests, with no allocation or lifetime.
pub type Sink = fn(&str);

static SINK: Mutex<Option<Sink>> = Mutex::new(None);

/// Install (`Some`) or clear (`None`) the process-wide progress sink.
///
/// The sink is global because parsing is: two concurrent parses share the
/// chatter. A sink that captures state cannot be installed — pass a `fn` and
/// keep the state behind a `static`.
pub fn set_sink(sink: Option<Sink>) {
    if let Ok(mut guard) = SINK.lock() {
        *guard = sink;
    }
}

/// Emit one progress line, if a sink is installed.
pub(crate) fn emit(line: &str) {
    let sink = SINK.lock().ok().and_then(|guard| *guard);
    if let Some(sink) = sink {
        sink(line);
    }
}

//! A once-only callback guard for the native dialog response handlers.
//!
//! Every native dialog in this module has the same shape: a single
//! `FnOnce`-style result callback that must run **exactly once** even though
//! two widget signals can fire it (the OK/Cancel response, and Enter in an
//! entry, which is a separate "activate" signal). Without the guard, Enter
//! followed by the dialog's own response would call the caller twice — e.g.
//! committing an edit twice, or opening two follow-up dialogs.
//!
//! The guard also owns the `F` so the callback can be taken out of the
//! `RefCell` on first use (an `FnOnce` cannot be called twice), and reports
//! whether *this* signal was the first so a caller can run its
//! close-then-report sequencing only once.
//!
//! It is `Rc`-shared rather than owned by the closure because the same guard
//! is handed to several widget signals; `Rc<OnceCallback>` (or a clone) is
//! what each `connect_*` closure captures.
//!
//! Its callers are the six dialog bodies in `dialogs.rs`, each gated
//! `#[cfg(any(feature = "gui", feature = "gui-core"))]`, so `gui/mod.rs` gates
//! this module on the same pair.
//!
//! Three gates, and each earlier one was wrong, which is worth writing down:
//!
//!   * `gui` + `test` -- the original. True while the dialog wiring was
//!     GTK-only, and false as soon as a backend without the `gui` feature grew
//!     a dialog: the module compiled while its users could not name it.
//!   * every widget backend -- fixed that, and overshot. `pancurses` routes
//!     `App::run` to `pnc_backend`, so it has no dialog bodies either, and the
//!     struct and both methods became dead code.
//!   * the `gui_backend`-runs set (`gui` / `gui-core` / `zork` / `wasm`) --
//!     overshot again, for the same reason one level down: `zork` and `wasm`
//!     do run `gui_backend`, but their dialog bodies are `#[cfg(gui)]`
//!     no-ops, so neither constructs an `OnceCallback` either.
//!
//! The predicate that is actually right is therefore the same one on the
//! callers: `gui` or `gui-core`. The test is on the *bodies* having a body, not
//! on the backend being able to run.
//!
//! The module is pure `Rc`/`RefCell` logic and reaches no backend, so the gate
//! is purely "is there a dialog body that can call it".

use std::cell::RefCell;
use std::rc::Rc;

/// A callback that runs at most once, sharing its "already ran" flag with
/// every widget signal that can fire it.
pub struct OnceCallback<F> {
    inner: RefCell<(Option<F>, bool)>,
}

impl<F> OnceCallback<F> {
    /// Wrap `f` in a fresh, never-yet-run guard.
    pub fn new(f: F) -> Rc<Self> {
        Rc::new(OnceCallback {
            inner: RefCell::new((Some(f), false)),
        })
    }

    /// Claim the callback and hand it to `run`, but only the first time.
    ///
    /// Returns `true` when this call won the race (and `run` was invoked),
    /// `false` when a previous signal already consumed it — so a caller can
    /// gate its close/report side effects on the return value.
    pub fn claim(&self, run: impl FnOnce(F) -> bool) -> bool {
        let mut g = self.inner.borrow_mut();
        if g.1 {
            return false;
        }
        g.1 = true;
        match g.0.take() {
            Some(f) => {
                run(f);
                true
            }
            None => false,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn runs_the_callback_once_and_reports_the_winner() {
        let cb = OnceCallback::new(|v: &mut Vec<i32>| v.push(7));
        let mut log = Vec::new();
        assert!(cb.claim(|f| {
            f(&mut log);
            true
        }));
        assert_eq!(log, vec![7]);
        // A second signal must not run the callback again.
        assert!(!cb.claim(|f| {
            f(&mut log);
            true
        }));
        assert_eq!(log, vec![7]);
    }
}

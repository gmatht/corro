//! Core application logic: domain state, file I/O, cursor management,
//! undo/redo, linked sources, export, and business operations.
//!
//! This module has zero dependencies on crossterm or ratatui. It can be
//! reused with any UI toolkit.

pub mod action;
pub mod linked_source;
pub mod state;

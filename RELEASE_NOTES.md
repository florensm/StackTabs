# StackTabs Release Notes

## v1.1.0 (2025-03-19)

### Bug Fixes

- **Host activation after close** — The host window is now only activated when closing a tab if other tabs still remain. Previously, closing the last tab would briefly activate the (soon-to-be-destroyed) empty host; this is now avoided.
- **Tab switcher double-fire** — Pressing `Ctrl+Shift+A` while the tab switcher is already open no longer schedules a second instance. Instead, it now closes the switcher (toggle behavior), preventing duplicate overlays and focus issues.

### Theme Improvements

- **Tab corner radius** — All built-in themes now support `TabCornerRadius` in the `[Layout]` section for rounded tab corners. Default is `5`; use `0` for sharp corners.
- Themes updated with appropriate corner radius values for a more polished look.

### Documentation

- `themes/README.md` — Added `TabCornerRadius` to the Layout section reference.

---

**Full Changelog**: https://github.com/florensm/StackTabs/compare/09c29e3...v1.1.0

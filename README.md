# StackTabs

[![AutoHotkey v2](https://img.shields.io/badge/AutoHotkey-v2-334455?logo=autohotkey)](https://www.autohotkey.com/)
[![Platform](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows)](https://www.microsoft.com/windows)

Turn cluttered windows into one clean tabbed window. AutoHotkey v2 script that collects matching windows into a single host with tabs.

## Requirements

- [AutoHotkey v2](https://www.autohotkey.com/) installed

## Setup

1. Copy `config.ini.example` to `config.ini` and set `Match1` (and optionally `Match2`, `Match3`, …) to part of your target window's title.
2. Run `StackTabs.ahk` (double-click or via AutoHotkey). If no match patterns are configured, the script will prompt you and open the config file.

## Features

- **Simple collector window:** One resizable host for all matches
- **Auto-refresh:** Windows added and removed automatically
- **Single active view:** One embedded window shown at a time
- **Lightweight tabs:** Click to switch between captured windows
- **Pop-out:** Extract a tab into its own window for side-by-side use; popped-out tabs stay independent
- **Merge back:** Combine a popped-out tab back into the main stack
- **Taskbar icon:** Uses the active tab's app icon with a badge to distinguish StackTabs from the original app

## Hotkeys

| Hotkey | Action |
|--------|--------|
| `Win+Shift+T` | Show/hide the collector window |
| `Win+Shift+D` | Dump discovery scan to `discovery.txt` (only when `DebugDiscovery=1`) |
| `Ctrl+Tab` | Next tab (when StackTabs is focused) |
| `Ctrl+Shift+Tab` | Previous tab (when StackTabs is focused) |
| `Ctrl+W` | Close active tab |
| `Ctrl+Shift+O` | Pop out active tab into a separate StackTabs window |
| `Ctrl+Shift+M` | Merge popped-out tab back into main stack (when pop-out window is focused) |

## Configuration

Edit `config.ini` (copy from `config.ini.example` if needed). The file is read on startup. Sections:

- **[General]** — Window matching (Match1, Match2, …) and timing (RefreshInterval, CaptureDelayMs, …)
- **[Layout]** — Window size, tab bar, behavior (ShowOnlyWhenTabs, UseCustomTitleBar)
- **[Theme]** — Theme file (e.g. `ThemeFile=everforest.ini`)
- **[TitleFilters]** — Regex patterns to shorten tab titles (Strip1, Strip2, …)

### General

| Variable | Default | Description |
|----------|---------|-------------|
| `Match1` / `Match2` / … | *(required)* | Window title substrings to match (case-insensitive). At least one must be configured; use `Match1=PowerShell`, `Match2=Notepad`, etc. in `config.ini`. Legacy: `WindowTitleMatch` as fallback when no MatchN is set. |
| `g_TargetExe` | `""` | Optional process filter, e.g. `"powershell.exe"` |
| `g_RefreshInterval` | `200` | How often to rescan for windows (ms) |
| `g_CaptureDelayMs` | `900` | How long a matching window must exist before it is embedded |
| `g_TabDisappearGraceMs` | `300` | Grace period before removing a tab whose window disappeared |
| `g_DebugDiscovery` | `false` | When `1`, log new candidates to `discovery.txt` and enable Win+Shift+D to dump discovery scan. |

**Title filters:** In `[TitleFilters]`, add `Strip1`, `Strip2`, etc. Each is a regex pattern removed from window titles before display. Example: `Strip1=^App - \s*` strips a leading "App - " prefix.

### Layout

| Variable | Default | Description |
|----------|---------|-------------|
| `g_HostWidth` / `g_HostHeight` | `1200` / `800` | Initial window size (overridden by saved session position on next launch) |
| `g_HostPadding` | `8` | Padding around content |
| `g_HeaderHeight` | `44` | Tab bar height |
| `g_TabHeight` | `30` | Tab button height |
| `g_TabSlotMax` | `50` | Maximum number of tabs shown (extra tabs are hidden) |
| `g_TabTitleMaxLen` | `60` | Max characters in tab label (truncated with ellipsis) |
| `g_TabPosition` | `"top"` | Tab bar position: `top` or `bottom` |
| `g_ShowOnlyWhenTabs` | `true` | When `1`, show host only when 1+ tabs; hide to tray when 0 (default). Set to `0` to always show host |
| `g_UseCustomTitleBar` | `false` | Use borderless custom title bar (set to `1` in INI to enable) |
| `g_TitleBarHeight` | `28` | Custom title bar height when enabled |

### Themes

Themes are `.ini` files in the `themes\` folder. Switch themes from the tray menu: right-click the StackTabs icon, **Theme**, choose a theme. The selection is saved to `config.ini` under `[Theme]` / `ThemeFile=`.

**Built-in themes** include `dark`, `light`, `high-contrast`, `everforest`, `rose-pine`, `one-dark`, `tokyo-night`, `dracula`, `catppuccin-mocha`, `spacious`, `compact`, and more.

**Custom themes:** Add your own `.ini` files to `themes\` or `themes\custom\`. Both folders are auto-discovered. Use **Open themes folder** from the tray menu to open the themes directory.

**Creating a custom theme:** Copy an existing theme (e.g. `dark.ini`), rename it, and edit the colors. See [themes/README.md](themes/README.md) for the full key reference.

## How It Works

StackTabs uses standard Win32 window APIs to reparent matching windows into a custom host:

1. **Discovery:** Enumerates top-level windows, filters by title and optional process, scores child windows to find the best content surface
2. **Pending delay:** New windows must exist for `g_CaptureDelayMs` before embedding (avoids capturing transient shells)
3. **Embedding:** Saves original parent/owner/styles, reparents via `SetParent`, adjusts styles for child behavior
4. **Rebind:** If the app recreates a window (new HWND), the script matches it back by stable ID and rebinds

**Tab ID** is `processName|rootOwner|normalizedTitle|contentClass` so the same logical window is recognized even if the HWND changes.

## Safety

- **No system modification:** Does not modify system files, registry, or other processes' memory
- **Reversible:** On exit, all embedded windows are restored to their original parent, position, and styles
- **User-level only:** No elevation or admin rights required
- **Standard APIs:** Uses the same Win32 APIs as window managers and accessibility tools

## Edge Cases

| Scenario | Mitigation |
|----------|------------|
| App recreates windows aggressively | Rebind logic; grace period before removing tabs |
| Modal dialogs / file pickers | Open as separate top-level windows (expected) |
| Unusual window hierarchy | Set `DebugDiscovery=1`, use Win+Shift+D to inspect; adjust `ScoreContentCandidate` if needed |
| Focus / keyboard shortcuts | May behave differently in embedded window depending on target app |

## Notes

- Some applications do not like being re-parented and may repaint poorly
- Switching tabs hides and shows embedded windows; repaint behaviour depends on the target application
- Chromium-based browsers (Chrome, Edge) actively fight re-parenting and are not reliably supported

## License

MIT License. See [LICENSE](LICENSE) for details.

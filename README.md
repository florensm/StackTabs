# StackTabs

A lightweight AutoHotkey v2 script that collects all matching windows into one host window and shows one embedded window at a time with a simple tab row at the top.

## Requirements

- [AutoHotkey v2](https://www.autohotkey.com/) installed

## Setup

1. Edit `StackTabs.ahk` and set `g_WindowTitleMatch` to part of your program's window title:

```autohotkey
; Examples:
g_WindowTitleMatch := "Powershell"        ; PowerShell windows
g_WindowTitleMatch := "Notepad"          ; Notepad windows
g_WindowTitleMatch := "Chrome"           ; Chrome browser
g_WindowTitleMatch := "Remote Desktop"   ; RDP sessions
g_WindowTitleMatch := "Excel"            ; Excel workbooks
```

2. Run the script (double-click or from AHK).

## Features

- **Simple collector window**  One resizable host window for all matches
- **Auto-refresh**  Windows are added and removed automatically
- **Single active view**  Only one embedded window is shown at a time
- **Lightweight tabs**  Click a tab button to switch between captured windows
- **Pop-out windows**  Extract a tab into its own StackTabs window for side-by-side comparison; popped-out tabs stay independent and are not re-captured
- **Merge back**  Combine a popped-out tab back into the main stack

## Hotkeys

| Hotkey | Action |
|--------|--------|
| `Win+Shift+T` | Show/hide the collector window |
| `Win+Shift+D` | Dump discovery scan to debug file |
| `Ctrl+Tab` | Next tab (when StackTabs is focused) |
| `Ctrl+Shift+Tab` | Previous tab (when StackTabs is focused) |
| `Ctrl+W` | Close active tab |
| `Ctrl+Shift+O` | Pop out active tab into a separate StackTabs window |
| `Ctrl+Shift+M` | Merge popped-out tab back into main stack (when pop-out window is focused) |

## Configuration

You can configure StackTabs by editing variables at the top of `StackTabs.ahk` or by using `StackTabs.ini` (copy `StackTabs.ini.example` to `StackTabs.ini` and edit). The INI file is read on startup and overrides script defaults.

### General

| Variable | Default | Description |
|----------|---------|-------------|
| `g_WindowTitleMatch` | `"Ticket details"` | Window title substring to match (case-insensitive) |
| `g_TargetExe` | `""` | Optional process filter, e.g. `"powershell.exe"` |
| `g_RefreshInterval` | `200` | How often to rescan for windows (ms) |
| `g_CaptureDelayMs` | `900` | How long a matching window must exist before it is embedded |
| `g_TabDisappearGraceMs` | `300` | Grace period before removing a tab whose window disappeared |
| `g_TitleFilterPattern` | `""` | Regex to strip from window titles before showing in tabs (leave blank to disable) |
| `g_TitleFilterReplace` | `""` | Replacement string for the regex match (empty = remove) |

### Layout

| Variable | Default | Description |
|----------|---------|-------------|
| `g_HostWidth` / `g_HostHeight` | `1200` / `800` | Initial window size |
| `g_HostPadding` | `8` | Padding around content |
| `g_HeaderHeight` | `44` | Tab bar height |
| `g_TabHeight` | `30` | Tab button height |
| `g_TabSlotMax` | `50` | Maximum number of tabs shown (extra tabs are hidden) |
| `g_TabTitleMaxLen` | `60` | Max characters in tab label (truncated with ellipsis) |
| `g_UseCustomTitleBar` | `false` | Use borderless custom title bar (set to `1` in INI to enable) |
| `g_TitleBarHeight` | `28` | Custom title bar height when enabled |

### Theme

| Variable | Default | Description |
|----------|---------|-------------|
| `g_ThemePreset` | `"dark"` | Preset: `dark`, `light`, or `high-contrast` |
| `g_ThemeBackground` | `"1E1E1E"` | Host window background (hex) |
| `g_ThemeTabBarBg` | `"252526"` | Tab bar background |
| `g_ThemeTabActiveBg` | `"2D7DFF"` | Active tab background |
| `g_ThemeTabInactiveBg` | `"30343B"` | Inactive tab background |
| `g_ThemeTabInactiveBgHover` | `"3C4049"` | Inactive tab hover background |
| `g_ThemeContentBorder` | `"404040"` | Content area border color |
| `g_ThemeFontName` | `"Segoe UI"` | Font family |

### Theme presets (colors applied)

| Preset | Background | Tab bar | Active tab | Inactive tab | Text |
|--------|------------|---------|------------|--------------|------|
| `dark` | 1E1E1E | 252526 | 2D7DFF (blue) | 30343B | FFFFFF / D8DEE9 |
| `light` | F3F3F3 | E8E8E8 | 0078D4 (blue) | E1E1E1 | 333333 |
| `high-contrast` | 000000 | 1A1A1A | FFFF00 (yellow) | 333333 | FFFFFF / 000000 |

## How It Works

StackTabs uses standard Win32 window APIs to reparent matching windows into a custom host:

1. **Discovery**  Enumerates top-level windows, filters by title and optional process, scores child windows to find the best content surface
2. **Pending delay**  New windows must exist for `g_CaptureDelayMs` before embedding (avoids capturing transient shells)
3. **Embedding**  Saves original parent/owner/styles, reparents via `SetParent`, adjusts styles for child behavior
4. **Rebind**  If the app recreates a window (new HWND), the script can match it back by stable ID and rebind

**Tab ID** is `processName|rootOwner|normalizedTitle|contentClass` so the same logical window is recognized even if the HWND changes.

## Safety

- **No system modification**  Does not modify system files, registry, or other processes' memory
- **Reversible**  On exit, all embedded windows are restored to their original parent, position, and styles
- **User-level only**  No elevation or admin rights required
- **Standard APIs**  Uses the same Win32 APIs as window managers and accessibility tools

## Edge Cases

| Scenario | Mitigation |
|----------|------------|
| App recreates windows aggressively | Rebind logic; grace period before removing tabs |
| Modal dialogs / file pickers | Open as separate top-level windows (expected) |
| Unusual window hierarchy | Use Win+Shift+D to inspect; adjust `ScoreContentCandidate` if needed |
| Focus / keyboard shortcuts | May behave differently in embedded window depending on target app |

## Future Ideas

- Mouse wheel over tab bar to switch tabs
- Middle-click tab to close; right-click context menu
- Remember window position and last active tab between sessions
- Tab overflow: scroll or dropdown when too many tabs
- Configurable hotkeys and capture delay
- Drag to reorder tabs

## Notes

- Some applications do not like being re-parented and may repaint poorly
- Switching tabs hides and shows embedded windows; repaint behavior depends on the target application

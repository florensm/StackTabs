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

- **Simple collector window** - One resizable host window for all matches
- **Auto-refresh** - Windows are added and removed automatically
- **Single active view** - Only one embedded window is shown at a time
- **Lightweight tabs** - Click a tab button to switch between captured windows
- **Pop-out windows** - Extract a tab into its own StackTabs window for side-by-side comparison; popped-out tabs stay independent and are not re-captured
- **Merge back** - Combine a popped-out tab back into the main stack

## Hotkeys

| Hotkey | Action |
|--------|--------|
| `Win+Shift+T` | Show/hide the collector window |
| `Win+Shift+R` | Restore all windows and exit (saves work before reload) |
| `Win+Shift+D` | Dump discovery scan to debug file |
| `Ctrl+Tab` | Next tab (when StackTabs is focused) |
| `Ctrl+Shift+Tab` | Previous tab (when StackTabs is focused) |
| `Ctrl+W` | Close active tab |
| `Ctrl+Shift+O` | Pop out active tab into a separate StackTabs window |
| `Ctrl+Shift+M` | Merge popped-out tab back into main stack (when pop-out window is focused) |

## Window restoration

When you stop or reload the script, all embedded windows are restored to their original position and state. Use **Win+Shift+R** before reloading to restore windows cleanly and avoid work loss.

> If the script crashes or is force-killed, embedded windows may be lost (Windows destroys child windows when the parent process exits). Use Win+Shift+R to restore before restarting when possible.

## Configuration

| Variable | Default | Description |
|----------|---------|-------------|
| `g_WindowTitleMatch` | `"Powershell"` | Window title substring to match (case-insensitive) |
| `g_TargetExe` | `""` | Optional process filter, e.g. `"powershell.exe"` |
| `g_RefreshInterval` | `500` | How often to rescan for windows (ms) |
| `g_CaptureDelayMs` | `900` | How long a matching window must exist before it is embedded |
| `g_TabDisappearGraceMs` | `300` | Grace period before removing a tab whose window disappeared |

## How It Works

StackTabs uses standard Win32 window APIs to reparent matching windows into a custom host:

1. **Discovery** - Enumerates top-level windows, filters by title and optional process, scores child windows to find the best content surface
2. **Pending delay** - New windows must exist for `g_CaptureDelayMs` before embedding (avoids capturing transient shells)
3. **Embedding** - Saves original parent/owner/styles, reparents via `SetParent`, adjusts styles for child behavior
4. **Rebind** - If the app recreates a window (new HWND), the script can match it back by stable ID and rebind

**Tab ID** is `processName|rootOwner|normalizedTitle|contentClass` so the same logical window is recognized even if the HWND changes.

## Safety

- **No system modification** - Does not modify system files, registry, or other processes' memory
- **Reversible** - On exit, all embedded windows are restored to their original parent, position, and styles
- **User-level only** - No elevation or admin rights required
- **Standard APIs** - Uses the same Win32 APIs as window managers and accessibility tools

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

# StackTabs

[![AutoHotkey v2](https://img.shields.io/badge/AutoHotkey-v2-334455?logo=autohotkey)](https://www.autohotkey.com/)
[![Platform](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows)](https://www.microsoft.com/windows)

Turn cluttered windows into one clean tabbed window. AutoHotkey v2 script that collects matching windows into a single host with tabs.

## Requirements

- [AutoHotkey v2](https://www.autohotkey.com/) installed

## Quick Start

1. Copy `config.ini.example` to `config.ini`
2. Set `Match1` (and optionally `Match2`, `Match3`, …) to part of your target window's title
3. Run `StackTabs.ahk` (double-click or via AutoHotkey)

If no match patterns are configured, the script will prompt you and open the config file.

---

## Hotkeys & Shortcuts

### Global (work anywhere)

| Hotkey | Action |
|--------|--------|
| `Win+Shift+T` | Show/hide the collector window |
| `Win+Shift+D` | Dump discovery scan to `discovery.txt` *(only when `DebugDiscovery=1`)* |
| `Ctrl+Shift+A` | Open **Tab Switcher** overlay (fuzzy search across all tabs) |

### When StackTabs host is focused

| Hotkey | Action |
|--------|--------|
| `Ctrl+Tab` | Next tab |
| `Ctrl+Shift+Tab` | Previous tab |
| `Ctrl+1` … `Ctrl+9` | Jump directly to tab 1–9 by position |
| `Ctrl+W` | Close active tab |
| `Ctrl+Shift+O` | Pop out active tab into a separate window |
| `Ctrl+Shift+M` | Merge popped-out tab back into main stack *(when pop-out window is focused)* |

### Tab Switcher overlay (Ctrl+Shift+A)

| Key | Action |
|-----|--------|
| Type | Filter tabs by title (case-insensitive) |
| `↑` `↓` `←` `→` | Navigate between visible tab cards |
| `Enter` | Switch to selected tab and close overlay |
| `Escape` | Close overlay without switching |

---

## Configuration Reference

Edit `config.ini` (copy from `config.ini.example` if needed). **Restart StackTabs** after saving changes.

### [General] — Window matching & timing

| Variable | Default | Description |
|----------|---------|-------------|
| `Match1` | *(required)* | Window title substring to match (case-insensitive). Window must contain at least one Match. |
| `Match2`, `Match3`, … | — | Additional match patterns. Use `Match2=Notepad`, `Match3=Visual Studio Code`, etc. |
| `WindowTitleMatch` | — | Legacy fallback when no `MatchN` is set. Prefer `Match1`. |
| `TargetExe` | `""` | Optional process filter, e.g. `pwsh.exe`. Leave empty to match any process. |
| `SlowSweepInterval` | `3000` | Fallback scan interval (ms). Shell Hook + WinEvent provide event-driven discovery; this catches edge cases. |
| `StackDelayMs` | `30` | Minimum wait before stacking (avoids ghost effects when windows appear very quickly). |
| `StackSwitchDelayMs` | `150` | Delay before switching to a newly stacked tab (lets content load to reduce glitch). |
| `WatchdogMaxMs` | `1500` | Max wait for hung windows before giving up. |
| `TabDisappearGraceMs` | `300` | Grace period before removing a tab whose window disappeared. |
| `DebugDiscovery` | `0` | When `1`, log new candidates to `discovery.txt` and enable `Win+Shift+D` to dump discovery scan. |

### [Layout] — Host window & tab bar

| Variable | Default | Description |
|----------|---------|-------------|
| `HostTitle` | `StackTabs` | Title shown in the host window. |
| `HostWidth` | `1200` | Initial window width (px). Overridden by saved session position on next launch. |
| `HostHeight` | `800` | Initial window height (px). |
| `HostMinWidth` | `700` | Minimum host width (px). |
| `HostMinHeight` | `500` | Minimum host height (px). |
| `HostPadding` | `8` | Padding around content (left, right, top) in pixels. |
| `HostPaddingBottom` | `-1` | Bottom padding (px). `-1` = use `HostPadding`. Set only when you want different bottom padding. |
| `HeaderHeight` | `36` | Tab bar height (px). |
| `TabHeight` | `30` | Tab button height (px). |
| `TabGap` | `6` | Gap between tabs (px). |
| `MinTabWidth` | `120` | Minimum tab width (px). |
| `MaxTabWidth` | `240` | Maximum tab width (px). |
| `TabSlotMax` | `50` | Maximum number of tabs shown; extra tabs are hidden. |
| `CloseButtonWidth` | `22` | Width of the close button on each tab (px). |
| `PopoutButtonWidth` | `22` | Width of the pop-out button on each tab (px). |
| `TabBarAlignment` | `center` | Vertical alignment of tabs within the bar: `top`, `center`, or `bottom`. |
| `TabBarOffsetY` | `-1` | Legacy pixel offset. `-1` = use `TabBarAlignment`. Prefer `TabBarAlignment`. |
| `TabPosition` | `top` | Tab bar position: `top` or `bottom`. |
| `TabIndicatorHeight` | `3` | Height of the active-tab indicator strip (px). `0` to disable. |
| `TabCornerRadius` | `5` | Tab corner radius (px). `0` for sharp corners. |
| `ActiveTabStyle` | `full` | `full` = active tab has distinct background; `indicator` = only accent strip, same bg as inactive. |
| `ShowOnlyWhenTabs` | `1` | `1` = hide to tray when no tabs, show when 1+ tabs. `0` = always show host. |
| `UseCustomTitleBar` | `0` | `1` = custom title bar that matches the theme. `0` = system title bar. |
| `TitleBarHeight` | `28` | Custom title bar height (px). Only relevant when `UseCustomTitleBar=1`. |

### [Layout] — Tab titles

| Variable | Default | Description |
|----------|---------|-------------|
| `TabTitleMaxLen` | *(omit)* | Optional. Omit for fully dynamic (fits tab width). Set to cap titles shorter (e.g. `60`). |
| `TabMaxLines` | `1` | `1` = single line with ellipsis; `2`+ = word-wrap to multiple lines. |
| `TabTitleAlignH` | `center` | Horizontal text alignment: `left`, `center`, `right`. |
| `TabTitleAlignV` | `center` | Vertical text alignment: `top`, `center`. Ignored when `TabMaxLines` ≥ 2. |
| `ShowTabNumbers` | `0` | `1` = prefix titles with position: `1. Title`, `2. Title`, etc. |

### [Theme]

| Variable | Default | Description |
|----------|---------|-------------|
| `ThemeFile` | `dark.ini` | Theme file from the `themes\` folder. |

**Built-in themes:** `dark`, `light`, `spacious`, `compact`, `spacious-dark`, `bottom-tabs`, `high-contrast`, `minimal`, `numbered`, `tmux`.

**Switch themes:** Right-click the StackTabs tray icon → **Theme** → choose a theme. The selection is saved to `config.ini`.

**Custom themes:** Add `.ini` files to `themes\` or `themes\custom\`. Use **Open themes folder** from the tray menu. See [themes/README.md](themes/README.md) for the full theme key reference.

### [TitleFilters]

| Variable | Description |
|----------|-------------|
| `Strip1`, `Strip2`, … | Regex patterns removed from window titles before display. Applied in order; result is trimmed. |

**Example:**
```ini
Strip1=^MyApp - \s*
Strip2=\s*- Microsoft Edge$
```

---

## Example config.ini

```ini
[General]
Match1=PowerShell
Match2=Notepad
TargetExe=
SlowSweepInterval=3000
StackDelayMs=100
StackSwitchDelayMs=150
WatchdogMaxMs=1500
TabDisappearGraceMs=300
DebugDiscovery=0

[Layout]
HostTitle=StackTabs
HostWidth=1200
HostHeight=800
HostMinWidth=700
HostMinHeight=500
HostPadding=8
HostPaddingBottom=-1
HeaderHeight=36
TabHeight=30
TabGap=6
MinTabWidth=120
MaxTabWidth=240
TabSlotMax=50
CloseButtonWidth=22
PopoutButtonWidth=22
TabBarAlignment=center
TabPosition=top
TabIndicatorHeight=3
TabCornerRadius=5
ActiveTabStyle=full
TabTitleMaxLen=
TabMaxLines=1
TabTitleAlignH=center
TabTitleAlignV=center
ShowTabNumbers=0
ShowOnlyWhenTabs=1
UseCustomTitleBar=0
TitleBarHeight=28

[Theme]
ThemeFile=dark.ini

[TitleFilters]
; Strip1=^App - \s*
; Strip2=\s*- Microsoft Edge$
```

---

## Features

- **Simple collector window** — One resizable host for all matching windows
- **Auto-refresh** — Windows added and removed automatically
- **Single active view** — One embedded window shown at a time
- **Lightweight tabs** — Click to switch between captured windows
- **Pop-out** — Extract a tab into its own window for side-by-side use
- **Merge back** — Combine a popped-out tab back into the main stack
- **Tab Switcher** — `Ctrl+Shift+A` for fuzzy search across all tabs
- **Taskbar icon** — Uses the active tab's app icon with a badge to distinguish StackTabs

---

## Tray Menu

Right-click the StackTabs icon in the system tray:

- **Theme** — Submenu to switch themes
- **Open themes folder** — Opens `themes\` in Explorer
- **Exit** — Quit StackTabs

---

## How It Works

1. **Discovery:** Shell Hook + WinEvent hooks detect new and destroyed windows. A slow sweep (every `SlowSweepInterval` ms) catches edge cases.
2. **Watchdog:** Windows wait at least `StackDelayMs` before stacking. Responsive windows stack once the delay has passed; hung windows are retried up to `WatchdogMaxMs`, then skipped.
3. **Embedding:** Saves original parent/owner/styles, reparents via `SetParent`, adjusts styles for child behavior.
4. **Rebind:** If the app recreates a window (new HWND), the script matches it back by stable ID and rebinds.

**Tab ID** is `processName|rootOwner|normalizedTitle|contentClass` so the same logical window is recognized even if the HWND changes.

---

## Safety

- **No system modification** — Does not modify system files, registry, or other processes' memory
- **Reversible** — On exit, all embedded windows are restored to their original parent, position, and styles
- **User-level only** — No elevation or admin rights required
- **Standard APIs** — Uses the same Win32 APIs as window managers and accessibility tools

---

## Edge Cases

| Scenario | Mitigation |
|----------|------------|
| App recreates windows aggressively | Rebind logic; grace period before removing tabs |
| Modal dialogs / file pickers | Open as separate top-level windows (expected) |
| Unusual window hierarchy | Set `DebugDiscovery=1`, use `Win+Shift+D` to inspect |
| Focus / keyboard shortcuts | May behave differently in embedded window depending on target app |

---

## Notes

- Some applications do not like being re-parented and may repaint poorly
- Switching tabs hides and shows embedded windows; repaint behaviour depends on the target application
- Chromium-based browsers (Chrome, Edge) actively fight re-parenting and are not reliably supported

---

## License

MIT License. See [LICENSE](LICENSE) for details.

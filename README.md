# StackTabs

[![AutoHotkey v2](https://img.shields.io/badge/AutoHotkey-v2-334455?logo=autohotkey)](https://www.autohotkey.com/)
[![Platform](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows)](https://www.microsoft.com/windows)

**StackTabs** is an [AutoHotkey v2](https://www.autohotkey.com/) script that **embeds matching top-level windows into one tabbed host**. You get a single resizable frame, a themed tab bar, and normal window controls—while each app still runs in its own process.

Good fit: terminals, dev tools, and many Win32 apps that tolerate reparenting. Poor fit: browsers and some UI frameworks that fight embedding (see **Notes**).

## Requirements

- [AutoHotkey v2](https://www.autohotkey.com/) installed

## Quick start

1. Copy `config.ini.example` to `config.ini`
2. Set `Match1` (and optionally `Match2`, `Match3`, …) to a substring of the window titles you want to collect
3. Run `StackTabs.ahk` (double-click or run with AutoHotkey)

If no match patterns are configured, the script prompts you and opens the config file. Only one instance runs at a time (`#SingleInstance Force`).

---

## Screenshots:
With StackTabs
<img width="3440" height="1440" alt="image" src="https://github.com/user-attachments/assets/c47aada2-dad5-48b4-adbe-2909b405a5f3" />

Without StackTabs
<img width="3440" height="1440" alt="image" src="https://github.com/user-attachments/assets/847c8976-c9d4-46f6-9a06-27438ba06979" />


## Hotkeys

### Global (anywhere)

| Hotkey | Action |
|--------|--------|
| `Win+Shift+T` | Show or hide the main host (respects **Show only when tabs**—won’t show an empty host) |
| `Win+Shift+D` | Write a discovery dump to `discovery.txt` *(only when `DebugDiscovery=1`)* |

### When the StackTabs host (or embedded content) is focused

| Hotkey | Action |
|--------|--------|
| `Ctrl+Tab` / `Ctrl+Shift+Tab` | Open the **tab switcher** and move forward/back through tabs |
| `Ctrl+1` … `Ctrl+9` | Jump to tab 1–9 by position |
| `Ctrl+W` | Close the active tab |
| `Ctrl+Shift+O` | Pop the active tab out to its own host window |
| `Ctrl+Shift+M` | Merge a focused **pop-out** host back into the main stack |

### Tab switcher (`Ctrl+Tab` style)

The switcher is the browser-like flow: **hold Ctrl**, press **Tab** / **Shift+Tab** to cycle (or **Tab** while the overlay is open). The selected tab’s content is **previewed live** behind the overlay.

| Input | Action |
|--------|--------|
| Release **Ctrl** | Commit to the highlighted tab and close the switcher |
| `Enter` | Same as commit (activate selection) |
| `Escape` | Cancel and restore the tab that was active when you opened the switcher |
| Arrow keys | Move selection (with live preview) |
| `J` / `K` | Move selection up/down (vim-style, Ctrl-tab mode only) |

With only one tab, `Ctrl+Tab` does not open the overlay (nothing to cycle).

---

## Configuration

Edit `config.ini` (copy from `config.ini.example` if needed). **Restart StackTabs** after saving.

### `[General]` — Matching and timing

| Variable | Default *(if key omitted)* | Description |
|----------|----------------------------|-------------|
| `Match1` | *(required in practice)* | Title substring (case-insensitive). The window must match at least one `MatchN`. |
| `Match2`, `Match3`, … | — | Extra patterns, e.g. `Match2=Notepad`. |
| `MatchN=` *(empty value)* | — | Special: matches windows with an empty/whitespace title—**always** pair with `TargetExe=` so you don’t grab everything untitled. |
| `WindowTitleMatch` | — | Legacy single pattern if no `MatchN` is set. Prefer `Match1`. |
| `TargetExe` | `""` | Optional process filter, e.g. `pwsh.exe`. Empty = any process. |
| `SlowSweepInterval` | `10000` | Periodic rescan (ms). Shell Hook + WinEvent handle most discovery; this catches stragglers. *(The example `config.ini.example` uses `3000`.)* |
| `StackDelayMs` | `30` | Minimum delay before stacking (reduces glitches when windows appear in quick succession). |
| `StackSwitchDelayMs` | `150` | Delay before auto-switching to a newly stacked tab (lets content settle). |
| `WatchdogMaxMs` | `1500` | Max time to wait on a stubborn window before skipping it. |
| `TabDisappearGraceMs` | `300` | Grace before removing a tab whose window vanished. |
| `DebugDiscovery` | `0` | `1` = log discovery to `discovery.txt` and enable `Win+Shift+D` dump. |
| `KeepAboveTabApps` | `0` | `1` (recommended) = continuously keep the StackTabs host above any visible unowned window of a tracked tab's process, while letting that process's popups/dialogs stay on top. Handles apps that periodically raise their main shell (via `BringWindowToTop` or an internal z-order change) without moving focus, as well as dialogs that would otherwise become buried below the host. Focus is never changed; only visual z-order is adjusted. |

### `[Layout]` — Host and tab bar

| Variable | Default | Description |
|----------|---------|-------------|
| `HostTitle` | `StackTabs` | Host window title (and tray label base). |
| `HostWidth` / `HostHeight` | `1200` / `800` | Initial size (px). Session geometry can override on next launch. |
| `HostMinWidth` / `HostMinHeight` | `700` / `500` | Minimum host size (px). |
| `HostPadding` | `8` | Padding around content (left, right, top). |
| `HostPaddingBottom` | `-1` | Bottom padding; `-1` = same as `HostPadding`. |
| `HeaderHeight` | `36` | Tab bar height (px). |
| `TabHeight` | `30` | Tab button height (px). |
| `TabGap` | `6` | Gap between tabs (px). |
| `MinTabWidth` / `MaxTabWidth` | `120` / `240` | Tab button width limits (px). |
| `TabSlotMax` | `50` | Max visible tabs; additional tabs are off-screen in the strip. |
| `CloseButtonWidth` / `PopoutButtonWidth` | `22` | Per-tab chrome (px). |
| `TabBarAlignment` | `center` | `top`, `center`, or `bottom` within the bar. |
| `TabBarOffsetY` | `-1` | Legacy offset; `-1` = use alignment. |
| `TabPosition` | `top` | `top` or `bottom`. |
| `TabIndicatorHeight` | `3` | Active indicator height (px); `0` = off. |
| `TabCornerRadius` | `5` | Rounded corners (px); `0` = square. |
| `ActiveTabStyle` | `full` | `full` = active tab has its own background; `indicator` = accent strip only. |
| `ShowOnlyWhenTabs` | `1` | `1` = hide host when there are zero tabs; `0` = always show host. |

### `[Layout]` — Tab titles

| Variable | Default | Description |
|----------|---------|-------------|
| `TabTitleMaxLen` | *(omit)* | Optional max label length; omit for width-based dynamic truncation. |
| `TabMaxLines` | `1` | `1` = one line + ellipsis; `2`+ = wrap. |
| `TabTitleAlignH` / `TabTitleAlignV` | `center` | Horizontal / vertical alignment. |
| `ShowTabNumbers` | `0` | `1` = prefix with `1.`, `2.`, … |

### `[Theme]`

| Variable | Default | Description |
|----------|---------|-------------|
| `ThemeFile` | `dark.ini` | File under `themes\` (see tray menu **Theme**). |

Built-in themes include `dark`, `light`, `spacious`, `compact`, `spacious-dark`, `bottom-tabs`, `high-contrast`, `minimal`, `numbered`, `tmux`. Custom `.ini` files go in `themes\` or `themes\custom\`. Details: [themes/README.md](themes/README.md).

### `[TitleFilters]`

| Variable | Description |
|----------|-------------|
| `Strip1`, `Strip2`, … | Regexes removed from titles before tabs/switcher (applied in order, then trimmed). |

**Example:**

```ini
Strip1=^MyApp - \s*
Strip2=\s*- Microsoft Edge$
```

---

## Example `config.ini` skeleton

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
KeepAboveTabApps=0

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

[Theme]
ThemeFile=dark.ini

[TitleFilters]
; Strip1=^App - \s*
```

---

## What you get

- **One host, many apps** — Reparent embedded clients into a shared client area; one tab visible at a time.
- **Event-driven discovery** — Shell Hook plus WinEvent (show, name change, uncloak) plus a slow sweep for edge cases.
- **Stable tab identity** — Tab IDs survive HWND churn so recreated windows can rebind to the same tab.
- **Pop-out / merge** — Move a tab to its own StackTabs window or pull it back.
- **Ctrl+Tab switcher** — Overlay with live preview; commit on Ctrl release or Enter.
- **Themes** — INI-driven colors, fonts, and layout hints from `themes\`.
- **Tray icon** — Reflects the active tab’s application icon where possible.

---

## Tray menu

Right-click the tray icon: **Theme**, **Open themes folder**, **Exit**.

---

## How it works (short)

1. **Discovery** — New/destroyed windows via Shell Hook and WinEvent; periodic `RefreshWindows` as a safety net.
2. **Watchdog** — Candidates wait `StackDelayMs` (and title stability) before embed; hung windows retry until `WatchdogMaxMs`.
3. **Embedding** — Saves original parent/styles, `SetParent` into the host client area, adjusts styles; optional top-level “shell” window hidden when it differs from the embedded client.
4. **Rebind** — If the app replaces its window, matching by tab ID reattaches logic without you reopening StackTabs.

Tab ID shape: `processName|rootOwner|normalizedTitle|contentClass` (see code comments for details).

---

## Safety

- **No system surgery** — No registry edits, no injection into other processes.
- **Reversible on exit** — Embedded windows are detached and restored toward their prior parent/styles.
- **User session** — No admin required; normal Win32 APIs only.

---

## Edge cases

| Situation | What to expect |
|-----------|----------------|
| Aggressive HWND recycling | Rebind + grace period before dropping a tab |
| Some dialogs / pickers | Separate top-level windows; covered by `KeepAboveTabApps=1` |
| Weird hierarchies | `DebugDiscovery=1` and `Win+Shift+D` |
| Focus in embedded apps | Cross-process focus uses `AttachThreadInput`; some apps still need a click |

---

## Notes

- Reparenting is **unsupported** for many apps by design; expect quirks with Chrome/Edge and some WPF/WinUI stacks.
- Tab switches show/hide embedded HWNDs; repaint quality depends on the guest app.
- Make sure to copy `config.ini.example` and rename it to `config.ini`
- Some apps may not work, this script was made with a specific WPF based app in mind. 

---

## License

MIT License. See [LICENSE](LICENSE).

# StackTabs

[![AutoHotkey v2](https://img.shields.io/badge/AutoHotkey-v2-334455?logo=autohotkey)](https://www.autohotkey.com/)
[![Platform](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows)](https://www.microsoft.com/windows)

**StackTabs** is an [AutoHotkey v2](https://www.autohotkey.com/) script that **embeds matching top-level windows into one tabbed host**. You get a single resizable frame, a themed tab bar, and normal window controls, while each application keeps running in its own process.

It works well for terminals, many Win32 tools, and some desktop stacks that tolerate reparenting. Browsers and some modern UI frameworks often fight embedding; expect quirks (see **Compatibility**).

## Requirements

- [AutoHotkey v2](https://www.autohotkey.com/) installed on Windows

## Quick start

1. Copy `config.ini.example` to `config.ini` in the same folder as `StackTabs.ahk`.
2. Under `[General]`, set at least `Match1=` to a **case-insensitive substring** that must appear in the title of windows you want to stack. Add `Match2`, `Match3`, and so on as needed.
3. Run `StackTabs.ahk` (double-click or launch with AutoHotkey).

If no match patterns are configured, StackTabs prompts you and opens `config.ini`. Only one instance runs at a time (`#SingleInstance Force`).

After editing `config.ini`, **restart StackTabs** (tray menu or exit and run again) so changes load.

---

## Screenshots

With StackTabs  
<img width="3440" height="1440" alt="image" src="https://github.com/user-attachments/assets/c47aada2-dad5-48b4-adbe-2909b405a5f3" />

Without StackTabs  
<img width="3440" height="1440" alt="image" src="https://github.com/user-attachments/assets/847c8976-c9d4-46f6-9a06-27438ba06979" />

---

## Hotkeys

### Global (anywhere)

| Hotkey | Action |
|--------|--------|
| `Win+Shift+T` | Toggle the main host: **hide** when that window is active, **show** when it is hidden (or not in the foreground). |
| `Win+Shift+D` | Append a discovery dump to `discovery.txt` *(only when `DebugDiscovery=1` in `[General]`)* |

### When the StackTabs host or embedded content is focused

| Hotkey | Action |
|--------|--------|
| `Ctrl+Tab` / `Ctrl+Shift+Tab` | Open the **tab switcher** and move forward or backward through tabs |
| `Ctrl+1` … `Ctrl+9` | Jump to tab 1–9 by position |
| `Ctrl+W` | Close the active tab |
| `Ctrl+Shift+O` | Pop the active tab out to its own host window |
| `Ctrl+Shift+M` | Merge a focused **pop-out** host back into the main stack |

### Tab switcher (`Ctrl+Tab` flow)

Hold **Ctrl**, press **Tab** or **Shift+Tab** to cycle. The selected tab’s content is **previewed** behind the overlay. With a single tab, `Ctrl+Tab` does not open the overlay.

| Input | Action |
|--------|--------|
| Release **Ctrl** | Commit to the highlighted tab and close the switcher |
| `Enter` | Commit (same as releasing Ctrl) |
| `Escape` | Cancel; in Ctrl+Tab mode, restore the tab that was active when the overlay opened |
| `Left` / `Up` | Move selection toward the previous card |
| `Right` / `Down` | Move selection toward the next card |
| `K` / `J` | In Ctrl+Tab mode only: move selection **up** (K) or **down** (J), with wrapping |

---

## Configuration

Authoritative examples and inline comments live in **`config.ini.example`**. The tables below match what **`LoadConfigFromIni`** reads in `StackTabs.ahk`. If a key is missing, the **built-in default** from the `Config` object at the top of the script applies (not necessarily the numbers in the example file).

Boolean options use **`0`** or **`1`**.

### `[General]` — Matching, timing, and debugging

| Key | Default *(if omitted)* | Description |
|-----|--------------------------|-------------|
| `Match1`, `Match2`, … | *(none)* | Title **substrings** (case-insensitive). A window must match **at least one** `MatchN` entry. Patterns are read in order until an empty key stops the list. |
| `WindowTitleMatch` | *(empty)* | **Legacy:** single pattern used only when no `MatchN` keys are present. Prefer `Match1`, `Match2`, … |
| `TargetExe` | *(empty)* | Optional executable name filter (for example `pwsh.exe`). Empty = any process. |
| `SlowSweepInterval` | `10000` | Slow **rescan** interval in ms (safety net alongside Shell Hook and WinEvent). The example config uses `3000`. |
| `StackDelayMs` | `30` | Minimum wait before stacking; works together with title-stability checks. |
| `StackSwitchDelayMs` | `150` | Delay before auto-switching to a newly stacked tab. |
| `WatchdogMaxMs` | `1500` | Maximum time to keep retrying a stubborn embed. |
| `TabDisappearGraceMs` | `300` | Grace period before removing a tab whose window disappeared. |
| `DebugDiscovery` | `0` | `1` = log discovery details to **`discovery.txt`** and enable **`Win+Shift+D`**. |
| `KeepAboveTabApps` | `0` | `1` = keep the host above the tracked app’s main shell where needed, reparent dialogs so they stay above the host, and return focus to the host after popups close. |
| `KeepAboveTabAppsDebug` | `0` | With `KeepAboveTabApps=1`, writes verbose z-order tracing to **`debug-zorder.log`**. Troubleshooting only; the file grows quickly. |

**Empty titles:** you can use a `MatchN` line whose value is empty or whitespace-only to match windows with **no real title**, but you **must** combine that with `TargetExe=` so StackTabs does not capture every untitled window on the desktop. See comments in `config.ini.example`.

### `[Layout]` — Host window, tab bar, and tab chrome

| Key | Default *(if omitted)* | Description |
|-----|--------------------------|-------------|
| `HostTitle` | `StackTabs` | Host window title (and tray name base). |
| `HostWidth` / `HostHeight` | `1200` / `800` | Initial size in pixels. |
| `HostMinWidth` / `HostMinHeight` | `700` / `500` | Minimum host size in pixels. |
| `HostPadding` | `8` | Inset in pixels: trims **left/right** of the tab strip and embedded client; also sets the **vertical gap** between the tab bar and the client (`GetEmbedRect` in `StackTabs.ahk`). |
| `HostPaddingBottom` | `-1` | Bottom padding; **`-1`** means “use `HostPadding`”. |
| `HeaderHeight` | `36` | Tab bar strip height in pixels. |
| `TabPosition` | `top` | `top` or `bottom`. |
| `TabBarAlignment` | `center` | Vertical placement of tabs inside the bar: `top`, `center`, or `bottom`. |
| `TabHeight` | `30` | Tab button height in pixels. |
| `TabGap` | `6` | Gap between tabs in pixels. |
| `MinTabWidth` / `MaxTabWidth` | `120` / `240` | Tab width limits in pixels. |
| `CloseButtonWidth` / `PopoutButtonWidth` | `22` | Width of the per-tab close and pop-out controls. |
| `TabIndicatorHeight` | `3` | Height of the active-tab indicator strip; **`0`** disables it. |
| `TabCornerRadius` | `5` | Corner radius in pixels; **`0`** for square tabs. |
| `TabSeparatorWidth` | `0` | Width of optional vertical separators between tabs; pair with `TabSeparatorColor` in the theme file. |
| `ActiveTabStyle` | `full` | `full` = active tab has its own background. `indicator` = shared background with an accent strip. |
| `TabTitleMaxLen` | `9999` | Character cap for tab labels. **`9999`** (or omit) behaves as “no fixed cap”; width-based truncation still applies. |
| `TabMaxLines` | `1` | `1` = single line with ellipsis; **`2`** or more enables wrapping (increase `TabHeight` so text fits). |
| `TabTitleAlignH` | `center` | `left`, `center`, or `right`. |
| `TabTitleAlignV` | `center` | `top`, `center`, or `bottom` for single-line tabs. When **`TabMaxLines` ≥ 2**, the renderer **forces top** alignment so wrapped titles are not clipped. |
| `ShowTabNumbers` | `0` | `1` = prefix titles with `1.`, `2.`, … |
| `ShowCloseButton` | `1` | `1` = show the per-tab close button. |
| `ShowPopoutButton` | `1` | `1` = show the pop-out (detach) button. |

Theme files under `themes\` can override many of these same layout keys for the active theme only; see [themes/README.md](themes/README.md).

### `[Theme]`

| Key | Default *(if omitted)* | Description |
|-----|--------------------------|-------------|
| `ThemeFile` | `dark.ini` | Filename under **`themes\`** (tray menu **Theme**). If the file is missing, StackTabs falls back to **`themes\dark.ini`**. |

Shipped presets: **`dark.ini`** (fallback), **`spacious.ini`**, **`tmux.ini`**, plus **`obsidian.ini`**, **`ink.ini`**, **`silk.ini`**, and **`clay.ini`**. Add your own `.ini` next to them or under **`themes\custom\`** ([themes/README.md](themes/README.md)).

### `[TitleFilters]`

| Key | Description |
|-----|-------------|
| `Strip1`, `Strip2`, … | Regular expressions removed from window titles **before** display in tabs and the switcher. Applied in order; the result is trimmed. |

Example:

```ini
Strip1=^MyApp - \s*
Strip2=\s*- Microsoft Edge$
```

---

## Minimal `config.ini` shape

This mirrors the sections in `config.ini.example`; adjust keys to taste.

```ini
[General]
Match1=PowerShell
TargetExe=
SlowSweepInterval=3000
StackDelayMs=100
StackSwitchDelayMs=150
WatchdogMaxMs=1500
TabDisappearGraceMs=300

[Layout]
HostTitle=StackTabs
HostWidth=1200
HostHeight=800
HostMinWidth=700
HostMinHeight=500
HostPadding=8
HeaderHeight=36
TabPosition=top
TabBarAlignment=center
TabHeight=30
TabGap=6
MinTabWidth=120
MaxTabWidth=240
CloseButtonWidth=22
PopoutButtonWidth=22
TabIndicatorHeight=3
TabCornerRadius=5
ActiveTabStyle=full
TabMaxLines=1
TabTitleAlignH=center
TabTitleAlignV=center
ShowTabNumbers=0

[Theme]
ThemeFile=dark.ini

[TitleFilters]
; Strip1=^App - \s*
```

---

## What you get

- **One host, many processes** — Matching top-level windows are reparented into a shared client area; one tab is visible at a time.
- **Event-driven discovery** — Shell Hook and WinEvent for show, rename, and uncloak, plus a slow sweep for edge cases.
- **Stable tab identity** — Tab IDs are derived from stable window attributes so HWND churn can rebind to the same logical tab (see below).
- **Pop-out and merge** — Detach a tab to its own host window or pull it back.
- **Ctrl+Tab switcher** — Overlay with live preview; commit on Ctrl release or Enter.
- **Themes** — Colors, fonts, and optional layout overrides from `themes\*.ini` ([themes/README.md](themes/README.md)).
- **Tray** — Theme picker, open themes folder, exit; the icon reflects the active tab’s application where possible.

---

## Tray menu

Right-click the tray icon: **Theme**, **Open themes folder**, **Exit**.

---

## How it works (short)

1. **Discovery** — Hooks and events notice new and destroyed windows; `RefreshWindows` runs on `SlowSweepInterval` as a backstop.
2. **Embed gate** — Candidates wait `StackDelayMs` (and pass title stability) before embed; hung windows retry until `WatchdogMaxMs`.
3. **Embedding** — Original parent and styles are saved, `SetParent` moves the window into the host client area, and styles are adjusted. Optional logic can hide a separate “shell” HWND when it is not the embedded client.
4. **Rebind** — If an app replaces its window, the same tab ID can attach to the new HWND without restarting StackTabs.

**Tab ID** (see `BuildCandidateId` in the script): `pid|rootOwnerHwnd|normalizedTitle|contentWindowClass`. The HWND of the embedded client is **not** part of the ID so the tab can survive reflows and replacements.

---

## Safety

- **No system surgery** — No registry writes and no injection into other processes.
- **Reversible on exit** — Embedded windows are detached and restored toward their prior parent and styles.
- **User session** — Normal Win32 APIs; admin is not required.

---

## Compatibility and caveats

- Reparenting is **unsupported** by many applications by design. Chromium-based browsers, some WPF stacks, and WinUI surfaces are common trouble spots.
- Tab switches show and hide embedded HWNDs; repaint quality depends on the guest application.
- Cross-process focus uses `AttachThreadInput`; a few guests still need an explicit click after focus changes.

---

## License

MIT License. See [LICENSE](LICENSE).

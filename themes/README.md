# StackTabs Themes

Theme files are `.ini` files. Place them in `themes\` or `themes\custom\`. Switch themes from the tray menu: right-click the StackTabs icon, **Theme**, choose a theme. Shipped examples include `dark.ini` (runtime fallback when `ThemeFile` is missing), `spacious.ini`, `tmux.ini`, and `obsidian.ini`, `ink.ini`, `silk.ini`, `clay.ini`.

## [Theme] Section

Colors are hex RGB without `#` (e.g. `1C1C2E`).

| Key | Default | Description |
|-----|---------|-------------|
| `Background` | `1C1C2E` | Main window background |
| `TabBarBg` | `13132A` | Tab bar background |
| `TabActiveBg` | `7B6CF6` | Active tab background |
| `TabActiveText` | `FFFFFF` | Active tab text color |
| `TabInactiveBg` | `252540` | Inactive tab background |
| `TabInactiveBgHover` | `30304E` | Inactive tab background on hover |
| `TabInactiveText` | `C5CDF0` | Inactive tab text color |
| `TabIndicatorColor` | *(TabActiveBg)* | Color of the active-tab indicator strip. Lets themes use a distinct accent. |
| `IconColor` | `6878B0` | Close and pop-out button icon color |
| `ContentBorder` | `35355A` | Border around embedded window area |
| `WindowText` | `E0E8FF` | General window text (e.g. title bar) |
| `TabSeparatorColor` | *(omit)* | Color of vertical tab separators when `TabSeparatorWidth` is non-zero in `[Layout]` |

### Fonts

| Key | Default | Description |
|-----|---------|-------------|
| `FontName` | `Segoe UI` | Main font |
| `FontNameTab` | `Segoe UI Semibold` | Tab label font |
| `FontSize` | `9` | Main font size (pt) |
| `IconFont` | *(auto)* | Icon font (Segoe Fluent Icons or Segoe MDL2 Assets). Leave empty to auto-detect. |
| `IconFontSize` | `16` | Icon font size (pt) |

## [Layout] Section (Optional)

Themes can include a `[Layout]` section to override spacing and tab dimensions. Values not specified fall back to `config.ini` or script defaults.

| Key | Default | Description |
|-----|---------|-------------|
| `HostPadding` | `8` | Padding around content (px) |
| `HostPaddingBottom` | `-1` | Bottom padding (px). `-1` = use `HostPadding` (same on all sides). Set only when you want different bottom padding. |
| `HeaderHeight` | `36` | Tab bar height (px); matches script default when not set in `config.ini` |
| `TabGap` | `6` | Gap between tabs (px) |
| `MinTabWidth` | `120` | Minimum tab width (px) |
| `MaxTabWidth` | `240` | Maximum tab width (px) |
| `TabHeight` | `30` | Tab button height (px) |
| `CloseButtonWidth` | `22` | Close button width (px) |
| `PopoutButtonWidth` | `22` | Pop-out button width (px) |
| `TabBarAlignment` | `center` | Tab alignment within bar: `top`, `center`, or `bottom` |
| `TabIndicatorHeight` | `3` | Active tab indicator strip height (px). Use `0` to disable. |
| `TabCornerRadius` | `5` | Tab corner radius (px). Use `0` for sharp corners. |
| `TabSeparatorWidth` | `0` | Width in px of vertical lines between tabs; pair with `TabSeparatorColor` in `[Theme]` |
| `ActiveTabStyle` | `full` | `full` = active tab has different background; `indicator` = only the strip, same bg as inactive |
| `TabPosition` | `top` | Tab bar position: `top` or `bottom` |
| `TabTitleMaxLen` | *(omit)* | Optional. Omit for fully dynamic (fits tab width, works for 1 or 2 lines). Set to cap titles shorter. |
| `TabMaxLines` | `1` | `1` = single line with ellipsis; `2` = word-wrap to second line. |
| `TabTitleAlignH` | `center` | Horizontal text alignment: `left`, `center`, `right` |
| `TabTitleAlignV` | `center` | Vertical text alignment: `top`, `center`, `bottom`. When **`TabMaxLines` ≥ 2**, the host forces **top** alignment when drawing wrapped titles. |
| `ShowTabNumbers` | `0` | `1` = prefix titles with position (e.g. "1. Title", "2. Title") |
| `ShowCloseButton` | `1` | `1` = show per-tab close control |
| `ShowPopoutButton` | `1` | `1` = show pop-out (detach) control |

## Example

```ini
; My theme

[Theme]
Background=1C1C2E
TabBarBg=13132A
TabActiveBg=7B6CF6
TabActiveText=FFFFFF
TabIndicatorColor=7B6CF6
TabInactiveBg=252540
TabInactiveBgHover=30304E
TabInactiveText=C5CDF0
IconColor=6878B0
ContentBorder=35355A
WindowText=E0E8FF
FontName=Segoe UI
FontNameTab=Segoe UI Semibold
FontSize=9
IconFontSize=16

[Layout]
HostPadding=10
HeaderHeight=48
TabGap=8
TabHeight=34
TabIndicatorHeight=4
ActiveTabStyle=indicator
```

## Creating a Theme

1. Copy an existing theme (e.g. `dark.ini`, `silk.ini`, `obsidian.ini`, or `tmux.ini`)
2. Rename the file (e.g. `my-theme.ini`)
3. Edit colors and optionally add `[Layout]` overrides
4. Restart StackTabs or switch themes from the tray menu

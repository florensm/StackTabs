; StackTabs - owner-aware embedded window host
; AutoHotkey v2

#Requires AutoHotkey v2.0

; SetWindowPos / ShowWindow flags (Win32 API)
SWP_NOACTIVATE   := 0x0010
SWP_NOZORDER     := 0x0004
SWP_FRAMECHANGED := 0x0020
SWP_SHOWWINDOW   := 0x0040
SW_HIDE          := 0
SW_SHOWNOACTIVATE := 4
SW_SHOW          := 5
SW_SHOWNA        := 8

; ============ CONFIGURATION ============
; Paths (relative to script directory)
g_ConfigPath     := A_ScriptDir "\config.ini"
g_ConfigExample  := A_ScriptDir "\config.ini.example"
g_ThemesDir      := A_ScriptDir "\themes"
g_DebugLogPath   := A_ScriptDir "\discovery.txt"
g_DebugDiscovery := false  ; when true, AppendDebugLog writes to discovery.txt on new candidates

; Window title patterns: Match1/Match2/... in config.ini. Window must contain at least one.
g_WindowTitleMatches := []

; Optional EXE filter. Leave blank to match any process.
g_TargetExe := ""

; Shell Hook + Slow Sweep: event-driven discovery; fallback scan interval.
g_SlowSweepInterval := 3000
g_StackDelayMs := 30        ; Minimum wait before stacking; title-stability check provides additional protection
g_StackSwitchDelayMs := 150  ; Delay before switching to newly stacked tab; lets content load to reduce glitch
g_WatchdogMaxMs := 1500
g_TabDisappearGraceMs := 300

; Host window defaults.
g_HostTitle := "StackTabs"
g_HostWidth := 1200
g_HostHeight := 800
g_HostMinWidth := 700
g_HostMinHeight := 500
g_HostPadding := 8
g_HostPaddingBottom := -1   ; -1 = use HostPadding; >=0 = use this for bottom padding
g_HeaderHeight := 36
g_TabGap := 6
g_MinTabWidth := 120
g_MaxTabWidth := 240
g_TabHeight := 30
g_TabSlotMax := 50
g_CloseButtonWidth := 22
g_PopoutButtonWidth := 22
g_TabBarAlignment := "center"  ; top, center, or bottom â€” tabs aligned within the tab bar
g_TabBarOffsetY := -1         ; legacy: -1 = use alignment; >=0 = use as pixel offset (overrides alignment)
g_TabPosition   := "top"    ; "top" or "bottom"
g_TabIndicatorHeight := 3  ; height in px of the active-tab indicator strip; 0 to disable
g_TabCornerRadius := 5
g_ActiveTabStyle := "full"  ; "full" = active tab has different bg; "indicator" = only indicator strip, same bg as inactive
g_ShowOnlyWhenTabs := true  ; when true, show host only when 1+ tabs; hide to tray when 0 (default). Set to 0 to always show host.

; === THEME (loaded from themes\ folder; dark.ini is the default and fallback) ===
g_ThemeTabIndicatorColor := ""   ; set by LoadThemeFromFile; defaults to TabActiveBg
g_ThemeIconFont        := ""   ; auto-detected at startup; override in theme file with IconFont=
g_ActiveThemeFile      := "dark.ini"   ; overridden by ThemeFile= in config.ini
g_UseCustomTitleBar    := false
g_TitleBarHeight       := 28

; Icon codepoints from Segoe Fluent Icons / Segoe MDL2 Assets (same PUA values)
g_IconClose  := Chr(0xe894)
g_IconPopout := Chr(0xE8A7)
g_IconMerge  := Chr(0xe944)

; === TITLE FILTERS ===
; Strip patterns loaded from [TitleFilters] Strip1/Strip2/... in StackTabs.ini.
; Each is a regex removed from the window title before it appears as a tab label.
g_TitleStripPatterns  := []
; Maximum characters shown in a tab label. 9999 = no cap (fully dynamic from tab width); set lower to force shorter titles.
g_TabTitleMaxLen      := 9999
; Max lines in tab label (1 = single line with ellipsis; 2+ = word wrap to multiple lines).
g_TabMaxLines         := 1
; Tab title alignment: H = left, center, right; V = top, center, bottom (center V = vertically centered in tab).
g_TabTitleAlignH      := "center"
g_TabTitleAlignV      := "center"
; When true, tab titles are prefixed with their 1-based position: "1. Title", "2. Title", etc.
g_ShowTabNumbers      := false
g_ShowCloseButton     := true
g_ShowPopoutButton    := true
g_TabSeparatorWidth   := 0
g_ThemeTabSeparatorColor := ""

; Loads config.ini into globals; migrates from StackTabs.ini if needed.
LoadConfigFromIni() {
    ; Migrate from StackTabs.ini if config.ini doesn't exist
    if !FileExist(g_ConfigPath) && FileExist(A_ScriptDir "\StackTabs.ini")
        FileCopy(A_ScriptDir "\StackTabs.ini", g_ConfigPath)
    if !FileExist(g_ConfigPath)
        return
    iniPath := g_ConfigPath
    ; Read theme first so it's applied even if the try block below throws (e.g. invalid Layout values)
    g_ActiveThemeFile := Trim(IniRead(iniPath, "Theme", "ThemeFile", "dark.ini"))
    global g_WindowTitleMatches, g_TargetExe, g_SlowSweepInterval, g_StackDelayMs, g_StackSwitchDelayMs, g_WatchdogMaxMs, g_TabDisappearGraceMs, g_DebugDiscovery
    global g_HostTitle, g_HostWidth, g_HostHeight, g_HostMinWidth, g_HostMinHeight
    global g_HostPadding, g_HostPaddingBottom, g_HeaderHeight, g_TabGap, g_MinTabWidth, g_MaxTabWidth, g_TabHeight
    global g_TabSlotMax, g_CloseButtonWidth, g_PopoutButtonWidth, g_TabBarAlignment, g_TabBarOffsetY, g_TabPosition
    global g_ShowOnlyWhenTabs, g_TabIndicatorHeight, g_ActiveTabStyle
    global g_TitleStripPatterns, g_TabTitleMaxLen, g_TabMaxLines, g_TabTitleAlignH, g_TabTitleAlignV, g_ShowTabNumbers
    global g_UseCustomTitleBar, g_TitleBarHeight
    global g_ActiveThemeFile
    try {
        g_TargetExe := IniRead(iniPath, "General", "TargetExe", g_TargetExe)
        g_SlowSweepInterval := Integer(IniRead(iniPath, "General", "SlowSweepInterval", g_SlowSweepInterval))
        g_StackDelayMs := Integer(IniRead(iniPath, "General", "StackDelayMs", g_StackDelayMs))
        g_StackSwitchDelayMs := Integer(IniRead(iniPath, "General", "StackSwitchDelayMs", g_StackSwitchDelayMs))
        g_WatchdogMaxMs := Integer(IniRead(iniPath, "General", "WatchdogMaxMs", g_WatchdogMaxMs))
        g_TabDisappearGraceMs := Integer(IniRead(iniPath, "General", "TabDisappearGraceMs", g_TabDisappearGraceMs))
        g_DebugDiscovery := (IniRead(iniPath, "General", "DebugDiscovery", "0") = "1")
        g_HostTitle := IniRead(iniPath, "Layout", "HostTitle", g_HostTitle)
        g_HostWidth := Integer(IniRead(iniPath, "Layout", "HostWidth", g_HostWidth))
        g_HostHeight := Integer(IniRead(iniPath, "Layout", "HostHeight", g_HostHeight))
        g_HostMinWidth := Integer(IniRead(iniPath, "Layout", "HostMinWidth", g_HostMinWidth))
        g_HostMinHeight := Integer(IniRead(iniPath, "Layout", "HostMinHeight", g_HostMinHeight))
        g_HostPadding := Integer(IniRead(iniPath, "Layout", "HostPadding", g_HostPadding))
        g_HostPaddingBottom := Integer(IniRead(iniPath, "Layout", "HostPaddingBottom", "-1"))
        g_HeaderHeight := Integer(IniRead(iniPath, "Layout", "HeaderHeight", g_HeaderHeight))
        g_TabGap := Integer(IniRead(iniPath, "Layout", "TabGap", g_TabGap))
        g_MinTabWidth := Integer(IniRead(iniPath, "Layout", "MinTabWidth", g_MinTabWidth))
        g_MaxTabWidth := Integer(IniRead(iniPath, "Layout", "MaxTabWidth", g_MaxTabWidth))
        g_TabHeight := Integer(IniRead(iniPath, "Layout", "TabHeight", g_TabHeight))
        g_TabSlotMax := Integer(IniRead(iniPath, "Layout", "TabSlotMax", g_TabSlotMax))
        g_CloseButtonWidth := Integer(IniRead(iniPath, "Layout", "CloseButtonWidth", g_CloseButtonWidth))
        g_PopoutButtonWidth := Integer(IniRead(iniPath, "Layout", "PopoutButtonWidth", g_PopoutButtonWidth))
        rawAlign := IniRead(iniPath, "Layout", "TabBarAlignment", "")
        g_TabBarAlignment := (rawAlign != "") ? Trim(rawAlign) : "center"
        g_TabBarOffsetY := Integer(IniRead(iniPath, "Layout", "TabBarOffsetY", "-1"))  ; legacy: -1 = use alignment
        g_TabTitleMaxLen := Integer(Trim(StrSplit(IniRead(iniPath, "Layout", "TabTitleMaxLen", "9999"), ";")[1]))
        g_TabMaxLines := Max(1, Integer(Trim(StrSplit(IniRead(iniPath, "Layout", "TabMaxLines", g_TabMaxLines), ";")[1])))
        g_TabTitleAlignH := Trim(StrSplit(IniRead(iniPath, "Layout", "TabTitleAlignH", g_TabTitleAlignH), ";")[1])
        g_TabTitleAlignV := Trim(StrSplit(IniRead(iniPath, "Layout", "TabTitleAlignV", g_TabTitleAlignV), ";")[1])
        g_ShowTabNumbers := (Trim(StrSplit(IniRead(iniPath, "Layout", "ShowTabNumbers", "0"), ";")[1]) = "1")
        g_ShowCloseButton := (Trim(StrSplit(IniRead(iniPath, "Layout", "ShowCloseButton", "1"), ";")[1]) = "1")
        g_ShowPopoutButton := (Trim(StrSplit(IniRead(iniPath, "Layout", "ShowPopoutButton", "1"), ";")[1]) = "1")
        g_TabPosition := IniRead(iniPath, "Layout", "TabPosition", "top")
        g_TabIndicatorHeight := Integer(IniRead(iniPath, "Layout", "TabIndicatorHeight", "3"))
        g_TabCornerRadius := Integer(IniRead(iniPath, "Layout", "TabCornerRadius", "5"))
        g_TabSeparatorWidth := Integer(IniRead(iniPath, "Layout", "TabSeparatorWidth", "0"))
        g_ActiveTabStyle := Trim(IniRead(iniPath, "Layout", "ActiveTabStyle", "full"))
        ; ShowOnlyWhenTabs: show host only when 1+ tabs; hide to tray when 0 (default). Fallback for old config keys.
        rawVal := IniRead(iniPath, "Layout", "ShowOnlyWhenTabs", IniRead(iniPath, "Layout", "KeepHostAlive", IniRead(iniPath, "Layout", "HideHostWhenEmpty", "1")))
        ; Strip inline comment (; ...) and trim so "1   ; comment" parses as 1
        g_ShowOnlyWhenTabs := (Trim(StrSplit(rawVal, ";")[1]) = "1")
        g_UseCustomTitleBar := (IniRead(iniPath, "Layout", "UseCustomTitleBar", "0") = "1")
        g_TitleBarHeight := Integer(IniRead(iniPath, "Layout", "TitleBarHeight", "28"))
    }
    ; Load match patterns from [General] Match1/Match2/...
    g_WindowTitleMatches := []
    i := 1
    loop {
        val := IniRead(iniPath, "General", "Match" i, "")
        if val = ""
            break
        g_WindowTitleMatches.Push(val)
        i++
    }
    ; No fallback: require Match1/Match2/... or WindowTitleMatch to be configured
    if g_WindowTitleMatches.Length = 0 {
        fallback := Trim(IniRead(iniPath, "General", "WindowTitleMatch", ""))
        if fallback != ""
            g_WindowTitleMatches.Push(fallback)
    }
    ; Load strip patterns from [TitleFilters] section (Strip1, Strip2, ...)
    g_TitleStripPatterns := []
    i := 1
    loop {
        val := IniRead(iniPath, "TitleFilters", "Strip" i, "")
        if val = ""
            break
        g_TitleStripPatterns.Push(val)
        i++
    }
}

; Loads theme colors and layout overrides from an .ini file; falls back to dark.ini if missing.
LoadThemeFromFile(themePath) {
    global g_ConfigPath, g_ThemeBackground, g_ThemeTabBarBg, g_ThemeTabActiveBg, g_ThemeTabActiveText, g_ThemeTabIndicatorColor
    global g_ThemeTabInactiveBg, g_ThemeTabInactiveBgHover, g_ThemeTabInactiveText, g_ThemeIconColor
    global g_ThemeContentBorder, g_ThemeWindowText, g_ThemeFontName, g_ThemeFontNameTab
    global g_ThemeFontSize, g_ThemeIconFont, g_ThemeIconFontSize
    global g_HostPadding, g_HostPaddingBottom, g_HeaderHeight, g_TabGap, g_MinTabWidth, g_MaxTabWidth, g_TabHeight
    global g_CloseButtonWidth, g_PopoutButtonWidth, g_TabBarAlignment, g_TabBarOffsetY, g_TabPosition, g_TabIndicatorHeight, g_TabCornerRadius, g_TabSeparatorWidth
    global g_ActiveTabStyle, g_TabMaxLines, g_TabTitleMaxLen, g_TabTitleAlignH, g_TabTitleAlignV, g_ShowTabNumbers
    global g_ShowCloseButton, g_ShowPopoutButton
    ; Fall back to dark.ini if the requested theme file doesn't exist
    if !FileExist(themePath)
        themePath := g_ThemesDir "\dark.ini"
    ; Use dark.ini values as fallbacks for missing keys (partial theme files)
    g_ThemeBackground         := IniRead(themePath, "Theme", "Background",           "1C1C2E")
    g_ThemeTabBarBg           := IniRead(themePath, "Theme", "TabBarBg",             "13132A")
    g_ThemeTabActiveBg        := IniRead(themePath, "Theme", "TabActiveBg",          "7B6CF6")
    g_ThemeTabActiveText      := IniRead(themePath, "Theme", "TabActiveText",        "FFFFFF")
    ; TabIndicatorColor: color of the active-tab indicator strip. Defaults to TabActiveBg.
    g_ThemeTabIndicatorColor   := IniRead(themePath, "Theme", "TabIndicatorColor",    g_ThemeTabActiveBg)
    g_ThemeTabInactiveBg      := IniRead(themePath, "Theme", "TabInactiveBg",        "252540")
    g_ThemeTabInactiveBgHover := IniRead(themePath, "Theme", "TabInactiveBgHover",   "30304E")
    g_ThemeTabInactiveText    := IniRead(themePath, "Theme", "TabInactiveText",      "C5CDF0")
    g_ThemeIconColor          := IniRead(themePath, "Theme", "IconColor",            "6878B0")
    g_ThemeContentBorder      := IniRead(themePath, "Theme", "ContentBorder",        "35355A")
    g_ThemeWindowText         := IniRead(themePath, "Theme", "WindowText",           "E0E8FF")
    g_ThemeFontName           := IniRead(themePath, "Theme", "FontName",             "Segoe UI")
    g_ThemeFontNameTab        := IniRead(themePath, "Theme", "FontNameTab",          "Segoe UI Semibold")
    g_ThemeFontSize           := Integer(IniRead(themePath, "Theme", "FontSize",     "9"))
    g_ThemeIconFont           := IniRead(themePath, "Theme", "IconFont",             "")
    g_ThemeIconFontSize       := Integer(IniRead(themePath, "Theme", "IconFontSize",   "16"))
    ; Optional layout overrides â€” only applied if the theme file includes a [Layout] section
    g_HostPadding        := Integer(IniRead(themePath, "Layout", "HostPadding",        String(g_HostPadding)))
    g_HostPaddingBottom  := Integer(IniRead(themePath, "Layout", "HostPaddingBottom",  String(g_HostPaddingBottom)))
    g_HeaderHeight       := Integer(IniRead(themePath, "Layout", "HeaderHeight",       String(g_HeaderHeight)))
    g_TabGap             := Integer(IniRead(themePath, "Layout", "TabGap",             String(g_TabGap)))
    g_MinTabWidth        := Integer(IniRead(themePath, "Layout", "MinTabWidth",        String(g_MinTabWidth)))
    g_MaxTabWidth        := Integer(IniRead(themePath, "Layout", "MaxTabWidth",        String(g_MaxTabWidth)))
    g_TabHeight          := Integer(IniRead(themePath, "Layout", "TabHeight",          String(g_TabHeight)))
    g_CloseButtonWidth   := Integer(IniRead(themePath, "Layout", "CloseButtonWidth",   String(g_CloseButtonWidth)))
    g_PopoutButtonWidth  := Integer(IniRead(themePath, "Layout", "PopoutButtonWidth",  String(g_PopoutButtonWidth)))
    ; TabBarAlignment: when theme doesn't specify, read from config so we don't carry over previous theme's value
    rawAlign := IniRead(themePath, "Layout", "TabBarAlignment", "")
    g_TabBarAlignment := (rawAlign != "") ? Trim(rawAlign) : IniRead(g_ConfigPath, "Layout", "TabBarAlignment", "center")
    rawOffset := IniRead(themePath, "Layout", "TabBarOffsetY", "")
    if rawOffset != ""
        g_TabBarOffsetY := Integer(rawOffset)
    g_TabIndicatorHeight := Integer(IniRead(themePath, "Layout", "TabIndicatorHeight", String(g_TabIndicatorHeight)))
    g_TabCornerRadius := Integer(IniRead(themePath, "Layout", "TabCornerRadius", String(g_TabCornerRadius)))
    g_TabSeparatorWidth := Integer(IniRead(themePath, "Layout", "TabSeparatorWidth", String(g_TabSeparatorWidth)))
    g_ThemeTabSeparatorColor := IniRead(themePath, "Theme", "TabSeparatorColor", "")
    ; TabPosition: when theme doesn't specify, read from config (not g_TabPosition) so we don't carry over
    ; the previous theme's value when switching themes
    rawPos := IniRead(themePath, "Layout", "TabPosition", "")
    g_TabPosition := (rawPos != "") ? Trim(rawPos) : IniRead(g_ConfigPath, "Layout", "TabPosition", "top")
    rawStyle := IniRead(themePath, "Layout", "ActiveTabStyle", "")
    g_ActiveTabStyle := (rawStyle != "") ? Trim(rawStyle) : "full"
    rawTabMaxLines := IniRead(themePath, "Layout", "TabMaxLines", "")
    if rawTabMaxLines != "" {
        parts := StrSplit(rawTabMaxLines, ";")
        if parts.Length
            g_TabMaxLines := Max(1, Integer(Trim(parts[1])))
    }
    rawTabTitleMaxLen := IniRead(themePath, "Layout", "TabTitleMaxLen", "")
    if rawTabTitleMaxLen != "" {
        parts := StrSplit(rawTabTitleMaxLen, ";")
        if parts.Length
            g_TabTitleMaxLen := Integer(Trim(parts[1]))
    }
    rawAlignH := IniRead(themePath, "Layout", "TabTitleAlignH", "")
    if rawAlignH != "" {
        parts := StrSplit(rawAlignH, ";")
        if parts.Length
            g_TabTitleAlignH := Trim(parts[1])
    }
    rawAlignV := IniRead(themePath, "Layout", "TabTitleAlignV", "")
    if rawAlignV != "" {
        parts := StrSplit(rawAlignV, ";")
        if parts.Length
            g_TabTitleAlignV := Trim(parts[1])
    }
    rawShowNums := IniRead(themePath, "Layout", "ShowTabNumbers", "")
    if rawShowNums != "" {
        parts := StrSplit(rawShowNums, ";")
        if parts.Length
            g_ShowTabNumbers := (Trim(parts[1]) = "1")
    } else {
        cfgVal := IniRead(g_ConfigPath, "Layout", "ShowTabNumbers", "0")
        parts := StrSplit(cfgVal, ";")
        g_ShowTabNumbers := (parts.Length && Trim(parts[1]) = "1")
    }
    rawShowClose := IniRead(themePath, "Layout", "ShowCloseButton", "")
    if rawShowClose != "" {
        parts := StrSplit(rawShowClose, ";")
        if parts.Length
            g_ShowCloseButton := (Trim(parts[1]) = "1")
    } else {
        cfgVal := IniRead(g_ConfigPath, "Layout", "ShowCloseButton", "1")
        parts := StrSplit(cfgVal, ";")
        g_ShowCloseButton := (parts.Length && Trim(parts[1]) = "1")
    }
    rawShowPopout := IniRead(themePath, "Layout", "ShowPopoutButton", "")
    if rawShowPopout != "" {
        parts := StrSplit(rawShowPopout, ";")
        if parts.Length
            g_ShowPopoutButton := (Trim(parts[1]) = "1")
    } else {
        cfgVal := IniRead(g_ConfigPath, "Layout", "ShowPopoutButton", "1")
        parts := StrSplit(cfgVal, ";")
        g_ShowPopoutButton := (parts.Length && Trim(parts[1]) = "1")
    }
}


; Auto-detects Segoe Fluent Icons or falls back to Segoe MDL2 Assets for icon glyphs.
DetectIconFont() {
    global g_ThemeIconFont
    if g_ThemeIconFont != ""
        return
    try {
        Loop Reg, "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts", "V" {
            if InStr(A_LoopRegName, "Segoe Fluent Icons") {
                g_ThemeIconFont := "Segoe Fluent Icons"
                return
            }
        }
    }
    g_ThemeIconFont := "Segoe MDL2 Assets"
}

; Builds tray menu with theme submenu, themes folder link, and Exit.
BuildTrayMenu() {
    global g_ActiveThemeFile, g_ThemesDir
    A_TrayMenu.Delete()
    themeSubMenu := Menu()
    themesDir := g_ThemesDir
    if DirExist(themesDir) {
        ; Built-in themes (themes\*.ini)
        Loop Files, themesDir "\*.ini" {
            fileName := A_LoopFileName
            displayName := ThemeDisplayName(fileName)
            themeSubMenu.Add(displayName, ThemeMenuHandler.Bind(fileName))
            if (Trim(fileName) = Trim(g_ActiveThemeFile))
                try themeSubMenu.Check(displayName)
        }
        ; Custom themes (themes\custom\*.ini)
        customDir := themesDir "\custom"
        if DirExist(customDir) {
            themeSubMenu.Add()
            Loop Files, customDir "\*.ini" {
                fileName := "custom\" A_LoopFileName
                displayName := ThemeDisplayName(A_LoopFileName)
                themeSubMenu.Add(displayName, ThemeMenuHandler.Bind(fileName))
                if (Trim(fileName) = Trim(g_ActiveThemeFile))
                    try themeSubMenu.Check(displayName)
            }
        }
    }
    A_TrayMenu.Add("Theme", themeSubMenu)
    A_TrayMenu.Add()
    A_TrayMenu.Add("Open themes folder", (*) => Run(g_ThemesDir))
    A_TrayMenu.Add()
    A_TrayMenu.Add("Exit", (*) => ExitApp())
}

; Converts theme filename (e.g. "dark-blue.ini") to display name ("Dark Blue").
ThemeDisplayName(fileName) {
    name := RegExReplace(fileName, "\.ini$", "")
    name := StrReplace(name, "-", " ")
    result := ""
    capitalize := true
    Loop Parse, name {
        ch := A_LoopField
        if ch = " " {
            result .= ch
            capitalize := true
        } else if capitalize {
            result .= StrUpper(ch)
            capitalize := false
        } else {
            result .= ch
        }
    }
    return result
}

; Tray menu callback: delegates to SwitchTheme.
ThemeMenuHandler(themeFileName, *) {
    SwitchTheme(themeFileName)
}

; Switches active theme, saves to config, reloads layout, and applies to all hosts.
SwitchTheme(themeFileName) {
    global g_ConfigPath, g_ConfigExample, g_ActiveThemeFile, g_ThemesDir
    if !FileExist(g_ConfigPath) {
        if FileExist(A_ScriptDir "\StackTabs.ini")
            FileCopy(A_ScriptDir "\StackTabs.ini", g_ConfigPath)
        else if FileExist(g_ConfigExample)
            FileCopy(g_ConfigExample, g_ConfigPath)
    }
    themePath := g_ThemesDir "\" themeFileName
    if !FileExist(themePath) {
        MsgBox("Theme file not found: " themePath, "StackTabs", "Icon!")
        return
    }
    IniWrite(themeFileName, g_ConfigPath, "Theme", "ThemeFile")
    g_ActiveThemeFile := themeFileName
    ; Reset layout from config so theme fallbacks use config values, not previous theme's
    LoadConfigFromIni()
    LoadThemeFromFile(themePath)
    DetectIconFont()
    BuildTrayMenu()
    for host in GetAllHosts() {
        ApplyThemeToHost(host)
    }
}

; Applies current theme colors and fonts to a host window and its tab controls.
ApplyThemeToHost(host) {
    global g_ThemeBackground, g_ThemeTabBarBg, g_ThemeContentBorder, g_ThemeTabIndicatorColor, g_ThemeTabActiveBg
    global g_ThemeIconColor, g_ThemeIconFontSize, g_ThemeIconFont, g_ThemeWindowText, g_ThemeFontName, g_ThemeFontSize
    global g_UseCustomTitleBar
    if !host || !host.gui || !IsWindowExists(host.hwnd)
        return
    host.gui.BackColor := g_ThemeBackground
    host.gui.SetFont("s" g_ThemeFontSize " c" g_ThemeWindowText, g_ThemeFontName)
    if host.HasProp("tabBarBg") && host.tabBarBg
        host.tabBarBg.Opt("Background0x" g_ThemeTabBarBg)
    if host.HasProp("contentBorderTop") && host.contentBorderTop {
        host.contentBorderTop.Opt("Background0x" g_ThemeContentBorder)
        host.contentBorderBottom.Opt("Background0x" g_ThemeContentBorder)
        host.contentBorderLeft.Opt("Background0x" g_ThemeContentBorder)
        host.contentBorderRight.Opt("Background0x" g_ThemeContentBorder)
    }
    if host.HasProp("titleBarBg") && host.titleBarBg {
        host.titleBarBg.Opt("Background0x" g_ThemeTabBarBg)
        host.titleCloseBtn.Opt("Background0x" g_ThemeTabBarBg " c" g_ThemeIconColor)
        if g_UseCustomTitleBar && host.hwnd {
            try {
                rgb := Integer("0x" g_ThemeTabBarBg)
                bgr := ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF)
                DllCall("dwmapi\DwmSetWindowAttribute", "ptr", host.hwnd, "int", 34, "uint*", &bgr, "uint", 4)
            }
        }
    }
    LayoutTabButtons(host)
    ShowOnlyActiveTab(host)
    DrawTabBar(host)
    UpdateHostTitle(host)
    RedrawAnyWindow(host.hwnd)
}

; ============ STATE ============
g_MainHost := ""             ; HostInstance for main window
g_PopoutHosts := []          ; array of HostInstance for popped-out windows
g_AllHostsCache := []
g_PendingCandidates := Map() ; tabId -> {firstSeen, candidate} (main host only)
g_WatchdogTimerActive := false
g_WatchdogInterval := 50
g_IsCleaningUp := false

LoadConfigFromIni()
if g_WindowTitleMatches.Length = 0 {
    cfgPath := FileExist(g_ConfigPath) ? g_ConfigPath : g_ConfigExample
    MsgBox("No window match patterns configured.`n`n"
        . "Add Match1=, Match2=, etc. in config.ini under [General].`n"
        . "Each value is a substring to match in window titles (e.g. Match1=PowerShell, Match2=Notepad).`n`n"
        . "See config.ini.example for details.", "StackTabs - Configuration Required", "Icon!")
    try Run(cfgPath)
    ExitApp()
}
LoadThemeFromFile(A_ScriptDir "\themes\" g_ActiveThemeFile)
DetectIconFont()
BuildTrayMenu()
if g_UseCustomTitleBar
    OnMessage(0x83, OnWmNcCalcSize)
DllCall("LoadLibrary", "Str", "gdiplus", "Ptr")
g_GdipToken := GdiplusStartup()
g_CachedFontFamily := Map()
g_CachedFont := Map()
g_CachedStringFormat := 0
g_GdipShutdownPending := false
BuildHostInstance(false)  ; create main host
; Shell Hook for event-driven window discovery
DllCall("RegisterShellHookWindow", "Ptr", g_MainHost.hwnd)
g_ShellHookMsg := DllCall("RegisterWindowMessage", "Str", "SHELLHOOK", "UInt")
OnMessage(g_ShellHookMsg, OnShellHook)

; WinEvent hooks: faster window detection than the Shell Hook for apps (e.g. WPF) that
; register with the taskbar late. One range registration covers the three events we care about:
;   EVENT_OBJECT_SHOW       (0x8002): window becomes visible — fires before taskbar registration
;   EVENT_OBJECT_NAMECHANGE (0x800C): title changes — catches WPF apps that set title after show
;   EVENT_OBJECT_UNCLOAKED  (0x8018): DWM uncloaks window — catches UWP/WinUI apps
; The callback filters to only those three; everything else in the range is discarded cheaply.
; WINEVENT_OUTOFCONTEXT (0): events are queued to this thread's message loop — no injection.
g_WinEventHookCallback := CallbackCreate(WinEventProc, , 7)
g_WinEventHooks := []
; EVENT_OBJECT_SHOW (0x8002): window becomes visible
g_WinEventHooks.Push(DllCall("SetWinEventHook",
    "UInt", 0x8002, "UInt", 0x8002,
    "Ptr",  0,      "Ptr",  g_WinEventHookCallback,
    "UInt", 0,      "UInt", 0, "UInt", 0, "Ptr"))
; EVENT_OBJECT_NAMECHANGE (0x800C): title changed (WPF sets title after show)
g_WinEventHooks.Push(DllCall("SetWinEventHook",
    "UInt", 0x800C, "UInt", 0x800C,
    "Ptr",  0,      "Ptr",  g_WinEventHookCallback,
    "UInt", 0,      "UInt", 0, "UInt", 0, "Ptr"))
; EVENT_OBJECT_UNCLOAKED (0x8018): UWP/WinUI window uncloaks
g_WinEventHooks.Push(DllCall("SetWinEventHook",
    "UInt", 0x8018, "UInt", 0x8018,
    "Ptr",  0,      "Ptr",  g_WinEventHookCallback,
    "UInt", 0,      "UInt", 0, "UInt", 0, "Ptr"))

OnExit(CleanupAll)

RefreshWindows()
SetTimer(RefreshWindows, g_SlowSweepInterval)

; ============ TAB SWITCHER OVERLAY ============
; Alt+Shift+F: floating overlay with tab cards and fuzzy search (when StackTabs is focused).
; Type to filter. Arrow keys navigate. Enter switches. Escape closes.

g_SwitcherGui       := ""
g_SwitcherAllTabs   := []
g_SwitcherVisible   := []
g_SwitcherSelVisIdx := 0
g_SwitcherCards     := []

; Tab switcher: Alt+Shift+F when StackTabs host is active
#HotIf StackTabsHostIsActive()
!+f:: {
    global g_SwitcherGui
    if g_SwitcherGui {
        SwitcherClose()
        return
    }
    ; Defer by 80ms so modifiers are physically released before the GUI opens.
    ; Without this the Edit control receives held modifiers = command mode, not typing mode.
    SetTimer(ShowTabSwitcher, -80)
}
#HotIf

ShowTabSwitcher() {
    global g_SwitcherGui, g_SwitcherAllTabs, g_SwitcherVisible, g_SwitcherSelVisIdx, g_SwitcherCards
    global g_ThemeTabBarBg, g_ThemeTabActiveBg, g_ThemeTabActiveText
    global g_ThemeTabInactiveBg, g_ThemeTabInactiveText, g_ThemeWindowText
    global g_ThemeFontName, g_ThemeFontNameTab, g_ThemeFontSize, g_ThemeBackground

    if g_SwitcherGui {
        SwitcherClose()
        return
    }

    allTabs := []
    for host in GetAllHosts() {
        for tabId in host.tabOrder {
            if !host.tabRecords.Has(tabId)
                continue
            record := host.tabRecords[tabId]
            if IsWindowExists(record.contentHwnd)
                allTabs.Push({tabId: tabId, host: host,
                    title: FilterTitle(record.title),
                    isActive: (tabId = host.activeTabId)})
        }
    }
    if allTabs.Length = 0
        return

    g_SwitcherAllTabs   := allTabs
    g_SwitcherVisible   := []
    g_SwitcherSelVisIdx := 1
    Loop allTabs.Length
        g_SwitcherVisible.Push(A_Index)
    for idx, item in allTabs {
        if item.isActive {
            g_SwitcherSelVisIdx := idx
            break
        }
    }

    ; Layout
    cardW   := 200
    cardH   := 56
    cardGap := 8
    pad     := 16
    searchH := 38
    cols    := Min(4, allTabs.Length)
    rows    := Ceil(allTabs.Length / cols)
    oW := Max(420, pad * 2 + cols * cardW + (cols - 1) * cardGap)
    oH := Min(A_ScreenHeight - 80,
              pad * 2 + searchH + cardGap + rows * cardH + (rows - 1) * cardGap)
    ox := (A_ScreenWidth  - oW) // 2
    oy := (A_ScreenHeight - oH) // 2

    g_SwitcherGui           := Gui("+AlwaysOnTop -Caption +ToolWindow", "TabSwitcher")
    g_SwitcherGui.BackColor  := g_ThemeTabBarBg
    g_SwitcherGui.MarginX    := 0
    g_SwitcherGui.MarginY    := 0
    g_SwitcherCards := []

    ; Search box — cue-banner text via EM_SETCUEBANNER after show
    g_SwitcherGui.SetFont("s" (g_ThemeFontSize + 1) " c" g_ThemeWindowText, g_ThemeFontName)
    searchBox := g_SwitcherGui.Add("Edit",
        "x" pad " y" pad " w" (oW - pad * 2) " h" searchH
        " Background" g_ThemeTabBarBg " c" g_ThemeWindowText, "")
    searchBox.SetFont("s" (g_ThemeFontSize + 1) " c" g_ThemeWindowText, g_ThemeFontName)

    cardAreaY := pad + searchH + cardGap

    Loop allTabs.Length {
        i    := A_Index
        col  := Mod(i - 1, cols)
        row  := (i - 1) // cols
        cx   := pad + col * (cardW + cardGap)
        cy   := cardAreaY + row * (cardH + cardGap)
        item := allTabs[i]
        isSel := (i = g_SwitcherSelVisIdx)
        bg := isSel ? g_ThemeTabActiveBg : g_ThemeTabInactiveBg
        fg := isSel ? g_ThemeTabActiveText : g_ThemeTabInactiveText

        card := g_SwitcherGui.Add("Text",
            "x" cx " y" cy " w" cardW " h" cardH
            " +0x200 Center Background" bg " c" fg,
            ShortTitle(item.title, 28))
        card.SetFont("s" g_ThemeFontSize " c" fg, g_ThemeFontNameTab)
        card.tabSwitcherIdx := i
        card.OnEvent("Click", SwitcherCardClick)
        g_SwitcherCards.Push(card)
    }

    g_SwitcherGui.OnEvent("Close", (*) => SwitcherClose())
    g_SwitcherGui.Show("x" ox " y" oy " w" oW " h" oH)
    DllCall("dwmapi.dll\DwmSetWindowAttribute",
        "ptr", g_SwitcherGui.Hwnd, "uint", 33, "uint*", 2, "uint", 4)

    ; Cue banner "Search tabs..." inside the edit box
    SendMessage(0x1501, 1, StrPtr("Search tabs..."),, "ahk_id " searchBox.Hwnd)

    searchBox.OnEvent("Change", SwitcherOnSearch)
    OnMessage(0x0100, SwitcherOnKeyDown, 1)
    WinActivate("ahk_id " g_SwitcherGui.Hwnd)
    searchBox.Focus()
}

SwitcherClose() {
    global g_SwitcherGui, g_SwitcherAllTabs, g_SwitcherVisible, g_SwitcherCards
    OnMessage(0x0100, SwitcherOnKeyDown, 0)
    if g_SwitcherGui {
        g_SwitcherGui.Destroy()
        g_SwitcherGui := ""
    }
    g_SwitcherAllTabs := []
    g_SwitcherVisible := []
    g_SwitcherCards   := []
}

SwitcherOnSearch(ctrl, *) {
    global g_SwitcherAllTabs, g_SwitcherVisible, g_SwitcherSelVisIdx
    query := ctrl.Value
    g_SwitcherVisible := []
    for idx, item in g_SwitcherAllTabs {
        if query = "" || InStr(item.title, query, false)
            g_SwitcherVisible.Push(idx)
    }
    g_SwitcherSelVisIdx := g_SwitcherVisible.Length ? 1 : 0
    SwitcherRefreshCards()
}

SwitcherRefreshCards() {
    global g_SwitcherGui, g_SwitcherVisible, g_SwitcherSelVisIdx, g_SwitcherCards
    global g_ThemeTabActiveBg, g_ThemeTabActiveText
    global g_ThemeTabInactiveBg, g_ThemeTabInactiveText
    global g_ThemeFontNameTab, g_ThemeFontSize
    if !g_SwitcherGui
        return

    visSet := Map()
    for _, tabIdx in g_SwitcherVisible
        visSet[tabIdx] := true
    selTabIdx := (g_SwitcherSelVisIdx >= 1 && g_SwitcherSelVisIdx <= g_SwitcherVisible.Length)
        ? g_SwitcherVisible[g_SwitcherSelVisIdx] : 0

    for idx, card in g_SwitcherCards {
        isVis := visSet.Has(idx)
        isSel := (idx = selTabIdx)
        card.Visible := isVis
        if isVis {
            bg := isSel ? g_ThemeTabActiveBg : g_ThemeTabInactiveBg
            fg := isSel ? g_ThemeTabActiveText : g_ThemeTabInactiveText
            card.Opt("Background" bg " c" fg)
            card.SetFont("s" g_ThemeFontSize " c" fg, g_ThemeFontNameTab)
        }
    }
}

SwitcherCardClick(ctrl, *) {
    global g_SwitcherVisible
    idx := ctrl.tabSwitcherIdx
    for _, tabIdx in g_SwitcherVisible {
        if tabIdx = idx {
            SwitcherActivate(tabIdx)
            return
        }
    }
}

SwitcherActivate(tabIdx) {
    global g_SwitcherAllTabs
    if tabIdx < 1 || tabIdx > g_SwitcherAllTabs.Length
        return
    item := g_SwitcherAllTabs[tabIdx]
    SwitcherClose()
    SelectTab(item.host, item.tabId)
    if item.host.hwnd && IsWindowExists(item.host.hwnd)
        WinActivate("ahk_id " item.host.hwnd)
}

SwitcherOnKeyDown(wParam, lParam, msg, hwnd) {
    global g_SwitcherGui, g_SwitcherSelVisIdx, g_SwitcherVisible, g_SwitcherAllTabs
    if !g_SwitcherGui
        return
    if hwnd != g_SwitcherGui.Hwnd {
        parent := DllCall("GetParent", "ptr", hwnd, "ptr")
        if parent != g_SwitcherGui.Hwnd
            return
    }
    cols  := Min(4, g_SwitcherAllTabs.Length)
    count := g_SwitcherVisible.Length
    if wParam = 0x1B {
        SwitcherClose()
        return 0
    }
    if count = 0
        return
    if wParam = 0x0D {
        if g_SwitcherSelVisIdx >= 1 && g_SwitcherSelVisIdx <= count
            SwitcherActivate(g_SwitcherVisible[g_SwitcherSelVisIdx])
        return 0
    }
    if wParam = 0x25 {
        g_SwitcherSelVisIdx := Max(1, g_SwitcherSelVisIdx - 1)
        SwitcherRefreshCards()
        return 0
    }
    if wParam = 0x27 {
        g_SwitcherSelVisIdx := Min(count, g_SwitcherSelVisIdx + 1)
        SwitcherRefreshCards()
        return 0
    }
    if wParam = 0x26 {
        g_SwitcherSelVisIdx := Max(1, g_SwitcherSelVisIdx - cols)
        SwitcherRefreshCards()
        return 0
    }
    if wParam = 0x28 {
        g_SwitcherSelVisIdx := Min(count, g_SwitcherSelVisIdx + cols)
        SwitcherRefreshCards()
        return 0
    }
}

; ============ HOTKEYS ============
; Win+Shift+T: toggle host visibility (hide when active, show when hidden).
#+t:: {
    global g_MainHost, g_ShowOnlyWhenTabs

    if !g_MainHost
        return

    if IsWindowExists(g_MainHost.hwnd) && WinActive("ahk_id " g_MainHost.hwnd)
        g_MainHost.gui.Hide()
    else {
        ; When ShowOnlyWhenTabs: only show if we have tabs
        if g_ShowOnlyWhenTabs && GetLiveTabCount(g_MainHost) = 0
            return
        g_MainHost.gui.Show()
    }
}

; Win+Shift+D: dump discovery debug to disk (only when DebugDiscovery=1).
#+d:: {
    global g_DebugDiscovery
    if !g_DebugDiscovery
        return
    DumpDiscoveryDebug()
}

#HotIf StackTabsHostIsActive()
^Tab:: {
    host := GetActiveStackTabsHost()
    if host
        CycleTabs(host, 1)
}
^+Tab:: {
    host := GetActiveStackTabsHost()
    if host
        CycleTabs(host, -1)
}
^w:: {
    host := GetActiveStackTabsHost()
    if host
        CloseActiveTab(host)
}
^+o:: {
    host := GetActiveStackTabsHost()
    if host && host.activeTabId && !host.isPopout
        PopOutTab(host, host.activeTabId)
}
^+m:: {
    host := GetActiveStackTabsHost()
    if host && host.isPopout && host.activeTabId
        MergeBackTab(host, host.activeTabId)
}
#HotIf

; Ctrl+1 through Ctrl+9: jump directly to tab by position (created via Hotkey so we can use a loop).
SelectTabByIndexHotkey(thisHotkey, *) {
    num := Integer(RegExReplace(thisHotkey, "\D", ""))
    if (host := GetActiveStackTabsHost()) && num >= 1 && num <= 9
        SelectTabByIndex(host, num)
}
HotIf StackTabsHostIsActive
Loop 9 {
    Hotkey "^" A_Index, SelectTabByIndexHotkey
}
HotIf

; Starts window drag when user clicks the custom title bar.
TitleBarDragClick(host, *) {
    MouseGetPos(&mx, &my)
    lParam := (mx & 0xFFFF) | (my << 16)
    PostMessage(0xA1, 2, lParam, , "ahk_id " host.hwnd)
}

; Closes popout host or hides main host when title bar close is clicked.
TitleBarCloseClick(host, *) {
    global g_PopoutHosts
    if host.isPopout {
        InvalidateHostsCache()
        for tabId in host.tabOrder.Clone()
            RemoveTrackedTab(host, tabId, true)
        for i, h in g_PopoutHosts {
            if h = host {
                g_PopoutHosts.RemoveAt(i)
                break
            }
        }
        if host.HasProp("iconHandle") && host.iconHandle {
            DllCall("DestroyIcon", "ptr", host.iconHandle)
            host.iconHandle := 0
        }
        host.gui.Destroy()
    } else {
        host.gui.Hide()
    }
}

; Returns true if the window handle is valid.
IsWindowExists(hwnd) {
    return !!DllCall("IsWindow", "ptr", hwnd, "int")
}

; Returns the count of live (existing) tabs in a host.
GetLiveTabCount(host) {
    if !host || !host.HasProp("tabOrder")
        return 0
    count := 0
    for tabId in host.tabOrder {
        if host.tabRecords.Has(tabId) && IsWindowExists(host.tabRecords[tabId].contentHwnd)
            count++
    }
    return count
}

; Gives keyboard focus to embedded content (from another process). Uses AttachThreadInput so
; SetFocus works cross-process. Without this, Ctrl+P etc. in the embedded app require a click first.
FocusEmbeddedContent(hostHwnd, contentHwnd) {
    if !hostHwnd || !contentHwnd || !IsWindowExists(contentHwnd)
        return
    targetTid := DllCall("GetWindowThreadProcessId", "ptr", contentHwnd, "ptr", 0, "uint")
    ourTid := DllCall("GetCurrentThreadId", "uint")
    if targetTid = ourTid
        return
    if !DllCall("AttachThreadInput", "uint", ourTid, "uint", targetTid, "int", 1)
        return
    try {
        DllCall("SetForegroundWindow", "ptr", hostHwnd)
        DllCall("SetFocus", "ptr", contentHwnd)
    } finally {
        DllCall("AttachThreadInput", "uint", ourTid, "uint", targetTid, "int", 0)
    }
}

; Sends WM_SYSCOMMAND SC_CLOSE and WM_CLOSE to reliably close a window (works with WPF).
CloseWindowReliably(topHwnd, contentHwnd := "") {
    if !contentHwnd
        contentHwnd := topHwnd
    prevHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    try {
        for hwnd in [topHwnd, contentHwnd] {
            if !hwnd || !IsWindowExists(hwnd)
                continue
            ; SendMessage (synchronous) - more reliable cross-process than PostMessage for WPF
            try SendMessage(0x0112, 0xF060, 0,, "ahk_id " hwnd,,,, 2000)  ; WM_SYSCOMMAND SC_CLOSE
            try SendMessage(0x0010, 0, 0,, "ahk_id " hwnd,,,, 2000)       ; WM_CLOSE
        }
    } finally {
        DetectHiddenWindows(prevHidden)
    }
}

; Returns the client-area width of a window.
GetClientWidth(hwnd) {
    try {
        WinGetClientPos(,, &w,, "ahk_id " hwnd)
        return w
    } catch {
        return 0
    }
}

; Returns the client-area height of a window.
GetClientHeight(hwnd) {
    try {
        WinGetClientPos(,,, &h, "ahk_id " hwnd)
        return h
    } catch {
        return 0
    }
}

InvalidateHostsCache() {
    global g_AllHostsCache
    g_AllHostsCache := []
}

; Returns array of main host plus all popout hosts.
GetAllHosts() {
    global g_MainHost, g_PopoutHosts, g_AllHostsCache
    if g_AllHostsCache.Length
        return g_AllHostsCache
    hosts := []
    if g_MainHost
        hosts.Push(g_MainHost)
    for h in g_PopoutHosts
        hosts.Push(h)
    g_AllHostsCache := hosts
    return hosts
}

; Finds the host that owns the given hwnd (or its parent chain).
GetHostForHwnd(hwnd) {
    current := hwnd
    while current {
        for host in GetAllHosts() {
            if host.hwnd = current
                return host
            for tabId, record in host.tabRecords {
                if (record.contentHwnd = current || record.topHwnd = current)
                    return host
            }
        }
        current := DllCall("GetParent", "ptr", current, "ptr")
    }
    return ""
}

; WM_NCCALCSIZE handler: extends client area into title bar to remove white bar (Windows 10/11).
OnWmNcCalcSize(hwnd, msg, lParam, wParam) {
    if !wParam || !lParam  ; wParam=1 means valid rects, lParam=struct pointer
        return
    for host in GetAllHosts() {
        if host.hwnd != hwnd
            continue
        ; Call DefWindowProc first; it modifies rgrc[0] to the client rect
        prevProc := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", -4, "ptr")
        result := DllCall("CallWindowProc", "ptr", prevProc, "ptr", hwnd, "uint", msg, "ptr", wParam, "ptr", lParam, "ptr")
        ; Get actual top border thickness (DPI-aware on Win 10 1703+)
        SM_CXPADDEDBORDER := 92
        SM_CYFRAME := 33
        dpi := 0
        try dpi := DllCall("GetDpiForWindow", "ptr", hwnd, "uint")
        pad := dpi ? DllCall("GetSystemMetricsForDpi", "int", SM_CXPADDEDBORDER, "uint", dpi, "int") : SysGet(SM_CXPADDEDBORDER)
        frameY := dpi ? DllCall("GetSystemMetricsForDpi", "int", SM_CYFRAME, "uint", dpi, "int") : SysGet(SM_CYFRAME)
        topBorder := frameY + pad
        top := NumGet(lParam, 4, "int")
        NumPut("int", top - topBorder, lParam, 4)
        return result
    }
}

; Creates a new host window (main or popout) with tab bar, content area, and theme.
BuildHostInstance(isPopout := false) {
    global g_MainHost, g_PopoutHosts
    global g_HostTitle, g_HostWidth, g_HostHeight, g_HostMinWidth, g_HostMinHeight
    global g_TabHeight, g_CloseButtonWidth
    global g_UseCustomTitleBar, g_TitleBarHeight
    global g_ShowOnlyWhenTabs

    host := Object()
    host.isPopout := isPopout
    host.tabRecords := Map()
    host.tabOrder := []
    host.activeTabId := ""
    host.tabHoveredId := ""
    host.tabScrollOffset := 0   ; index of first visible tab (0-based)
    host.tabScrollMax := 0      ; updated by DrawTabBar
    host.isResizing := false

    global g_ThemeBackground, g_ThemeWindowText, g_ThemeFontName, g_ThemeFontSize
    title := isPopout ? (g_HostTitle " (popped out)") : g_HostTitle
    guiOpts := "+Resize +MinSize" g_HostMinWidth "x" g_HostMinHeight
    if g_UseCustomTitleBar
        guiOpts .= " -Caption +Border"
    host.gui := Gui(guiOpts, title)
    host.gui.BackColor := g_ThemeBackground
    host.gui.MarginX := 0
    host.gui.MarginY := 0
    host.gui.SetFont("s" g_ThemeFontSize " c" g_ThemeWindowText, g_ThemeFontName)
    host.gui.OnEvent("Close", HostGuiClosed.Bind(host))
    host.gui.OnEvent("Size", HostGuiResized.Bind(host))

    global g_ThemeTabBarBg, g_ThemeContentBorder, g_HeaderHeight
    global g_ThemeIconFont, g_ThemeIconFontSize, g_IconClose
    tabBarY := 0
    tabBarH := g_HeaderHeight
    if g_UseCustomTitleBar {
        tabBarY := g_TitleBarHeight
        host.titleBarBg := host.gui.Add("Text", "x0 y0 w" g_HostWidth " h" g_TitleBarHeight " +0x200 +0x100 Background" g_ThemeTabBarBg, "")
        host.titleBarBg.OnEvent("Click", TitleBarDragClick.Bind(host))
        host.titleText := host.gui.Add("Text", "x8 y0 w" (g_HostWidth - 60) " h" g_TitleBarHeight " +0x200 +0x100 BackgroundTrans", title)
        host.titleText.OnEvent("Click", TitleBarDragClick.Bind(host))
        host.titleCloseBtn := host.gui.Add("Text", "x" (g_HostWidth - 46) " y0 w46 h" g_TitleBarHeight " +0x200 +0x100 Center Background" g_ThemeTabBarBg, g_IconClose)
        host.titleCloseBtn.SetFont("s" g_ThemeIconFontSize, g_ThemeIconFont)
        host.titleCloseBtn.Opt("c" g_ThemeIconColor)
        host.titleCloseBtn.OnEvent("Click", TitleBarCloseClick.Bind(host))
    }
    host.tabBarBg := host.gui.Add("Text", "x0 y" tabBarY " w" g_HostWidth " h" tabBarH " Background" g_ThemeTabBarBg, "")
    host.tabCanvas := host.gui.Add("Pic",
        "x0 y" tabBarY " w" g_HostWidth " h" tabBarH " +0xE", "")
    host.contentBorderTop := host.gui.Add("Text", "Hidden x0 y0 w0 h1 Background" g_ThemeContentBorder, "")
    host.contentBorderBottom := host.gui.Add("Text", "Hidden x0 y0 w0 h1 Background" g_ThemeContentBorder, "")
    host.contentBorderLeft := host.gui.Add("Text", "Hidden x0 y0 w1 h0 Background" g_ThemeContentBorder, "")
    host.contentBorderRight := host.gui.Add("Text", "Hidden x0 y0 w1 h0 Background" g_ThemeContentBorder, "")
    host.hwnd := host.gui.Hwnd
    host.clientHwnd := host.hwnd
    global g_ShowOnlyWhenTabs
    showOpts := "w" g_HostWidth " h" g_HostHeight
    ; Always pass Hide when the window should not be visible yet.
    ; Show() followed immediately by Hide() briefly makes the window visible to the OS,
    ; which causes tiling window managers to tile it as a floating window on first use.
    ; Passing Hide to Show() sets the size without ever showing the window.
    if isPopout || (!isPopout && g_ShowOnlyWhenTabs)
        showOpts .= " Hide"
    host.gui.Show(showOpts)

    ; Request Windows 11 rounded corners (no-op if already applied by system)
    cornerPref := 2  ; DWM_WCP_ROUND
    DllCall("dwmapi.dll\DwmSetWindowAttribute", "ptr", host.hwnd, "uint", 33, "uint*", cornerPref, "uint", 4)

    ; Match DWM window border to theme (fixes white/accent bar on Windows 11)
    if g_UseCustomTitleBar && host.hwnd {
        try {
            rgb := Integer("0x" g_ThemeTabBarBg)
            bgr := ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF)
            DllCall("dwmapi\DwmSetWindowAttribute", "ptr", host.hwnd, "int", 34, "uint*", &bgr, "uint", 4)
        }
    }

    if isPopout
        g_PopoutHosts.Push(host)
    else
        g_MainHost := host
    InvalidateHostsCache()
    return host
}

; Handles host close: restores tabs to main (popout) or hides main host (keeps script running).
HostGuiClosed(host, *) {
    global g_MainHost, g_PopoutHosts
    if host.isPopout {
        ; Restore all tabs in this popout to their original parent (release, don't merge)
        for tabId in host.tabOrder.Clone()
            RemoveTrackedTab(host, tabId, true)
        for i, h in g_PopoutHosts {
            if h = host {
                g_PopoutHosts.RemoveAt(i)
                break
            }
        }
        InvalidateHostsCache()
        if host.HasProp("iconHandle") && host.iconHandle
            DllCall("DestroyIcon", "ptr", host.iconHandle)
        host.gui.Destroy()
    } else {
        host.gui.Hide()
        return true  ; Prevent default close (keep script running in tray)
    }
}

; On host resize: re-layout tabs, update content area.
HostGuiResized(host, guiObj, minMax, width, height) {
    if minMax = -1
        return
    if host.HasProp("isResizing") && host.isResizing
        return
    if width < 100 || height < 100
        return
    host.isResizing := true
    try {
        LayoutTabButtons(host, width, height)
        ShowOnlyActiveTab(host)
    } finally {
        host.isResizing := false
    }
}

; Returns true if window has a non-empty title and is not hung.
IsReadyToStack(hwnd) {
    if !hwnd || !IsWindowExists(hwnd)
        return false
    title := SafeWinGetTitle(hwnd)
    if (title = "")
        return false
    hung := DllCall("User32.dll\IsHungAppWindow", "Ptr", hwnd, "Int")
    return !hung
}

; Adds candidate to pending map and starts watchdog timer; stacks after delay when title is stable.
TryStackOrPending(host, candidate) {
    global g_MainHost, g_PendingCandidates, g_StackDelayMs, g_WatchdogMaxMs, g_ShowOnlyWhenTabs, g_WatchdogTimerActive, g_WatchdogInterval
    if !g_PendingCandidates.Has(candidate.id) {
        ; First time seeing this candidate â€” record the timestamp
        now := A_TickCount
        g_PendingCandidates[candidate.id] := {firstSeen: now, candidate: candidate}
        AppendDebugLog("New candidate (pending watchdog)`r`n" candidate.hierarchySummary "`r`n")
    } else {
        ; Already pending â€” refresh metadata only, preserve firstSeen so the delay is not reset
        g_PendingCandidates[candidate.id].candidate := candidate
    }
    if !g_WatchdogTimerActive {
        g_WatchdogTimerActive := true
        g_WatchdogInterval := 50
        SetTimer(WatchdogCheck, -g_WatchdogInterval)
    }
}

; Processes a single pending candidate: builds fresh candidate, handles duplicate detection/dialog, creates tracked tab.
; Returns the new tab id on success, or "" if skipped/failed.
ProcessPendingCandidate(tabId, pending) {
    global g_MainHost
    ; Re-build candidate from scratch so we use fresh, stable metadata (not the snapshot from creation time)
    freshCandidate := BuildCandidateFromTopWindow(pending.candidate.topHwnd)
    if !IsObject(freshCandidate)
        return ""
    ; Skip if already embedded (could have been stacked by slow sweep between checks)
    if g_MainHost.tabRecords.Has(freshCandidate.id)
        return ""
    ; Duplicate detection: same process + same title = offer to close the older one
    existingTabId := FindTabWithSameTitle(g_MainHost, freshCandidate)
    if existingTabId != "" {
        result := DuplicateConfirmDialog(ShortTitle(freshCandidate.title, 50))
        if (result = "Yes")
            CloseTab(g_MainHost, existingTabId)
    }
    if CreateTrackedTab(g_MainHost, freshCandidate)
        return freshCandidate.id
    return ""
}

; Timer callback: stacks candidates that passed delay + title-stability; removes stale pending.
WatchdogCheck(*) {
    global g_MainHost, g_PendingCandidates, g_StackDelayMs, g_StackSwitchDelayMs, g_WatchdogMaxMs, g_WatchdogTimerActive, g_WatchdogInterval, g_ShowOnlyWhenTabs, g_DebugDiscovery
    if !g_MainHost || g_PendingCandidates.Count = 0 {
        g_WatchdogTimerActive := false
        SetTimer(WatchdogCheck, 0)
        g_WatchdogInterval := 50   ; reset for next activation
        return
    }
    now := A_TickCount
    toStack := []
    toRemove := []
    for tabId, pending in g_PendingCandidates {
        elapsed := now - pending.firstSeen
        if elapsed >= g_StackDelayMs && IsReadyToStack(pending.candidate.topHwnd) {
            ; Title-stability check: stack only once the title has been unchanged for two consecutive ticks (~50ms)
            currentTitle := SafeWinGetTitle(pending.candidate.topHwnd)
            if pending.HasProp("lastSeenTitle") && (currentTitle = pending.lastSeenTitle)
                toStack.Push({tabId: tabId, candidate: pending.candidate, firstSeen: pending.firstSeen})
            else {
                pending.lastSeenTitle := currentTitle
            }
        } else if elapsed >= g_WatchdogMaxMs {
            toRemove.Push(tabId)
        }
    }
    anyStacked := false
    lastStackedTabId := ""
    hadTabsBefore := g_MainHost.tabOrder.Length
    for item in toStack {
        g_PendingCandidates.Delete(item.tabId)
        if (newTabId := ProcessPendingCandidate(item.tabId, item)) {
            lastStackedTabId := newTabId
            anyStacked := true
        }
    }
    ; Defer GUI updates once for all stacked windows â€” avoids re-entrancy when gui.Add
    ; pumps messages mid-loop and HostGuiResized re-enters LayoutTabButtons.
    if anyStacked {
        if hadTabsBefore > 0 {
            SetTimer(WatchdogPostStackUpdate, -1)
            ; Cancel any previous pending switch timer before scheduling a new one.
            ; Each .Bind() produces a new object so we must store the reference to cancel it.
            if g_MainHost.HasProp("pendingSwitchTimer") && g_MainHost.pendingSwitchTimer
                SetTimer(g_MainHost.pendingSwitchTimer, 0)
            g_MainHost.pendingSwitchTimer := SwitchToNewTabDelayed.Bind(g_MainHost, lastStackedTabId)
            SetTimer(g_MainHost.pendingSwitchTimer, -g_StackSwitchDelayMs)
        } else {
            ; First tab: let the window settle hidden, then show. One-shot.
            g_MainHost.activeTabId := lastStackedTabId
            SetTimer(WatchdogPostStackUpdate, -g_StackSwitchDelayMs)
        }
    }
    for tabId in toRemove
        g_PendingCandidates.Delete(tabId)
    if g_PendingCandidates.Count = 0 {
        g_WatchdogTimerActive := false
        SetTimer(WatchdogCheck, 0)
        g_WatchdogInterval := 50
    } else {
        ; Adaptive back-off: start at 50ms, double each tick up to 400ms.
        ; Resets to 50ms when candidates are cleared.
        g_WatchdogInterval := Min(400, g_WatchdogInterval * 2)
        SetTimer(WatchdogCheck, -g_WatchdogInterval)
    }
}

; Deferred update after stacking: show host if needed, layout tabs, refresh content.
WatchdogPostStackUpdate(*) {
    global g_MainHost, g_ShowOnlyWhenTabs
    if !g_MainHost
        return
    ; Only call Show when the window is actually hidden â€” avoids repositioning an already-visible window
    if g_ShowOnlyWhenTabs && g_MainHost.tabOrder.Length >= 1 {
        if !WinExist("ahk_id " g_MainHost.hwnd)
            g_MainHost.gui.Show()  ; Show() activates by default; intentional for first appearance
        ; If already visible don't force-activate — user may have focus elsewhere
    }
    LayoutTabButtons(g_MainHost)
    ShowOnlyActiveTab(g_MainHost)
    UpdateHostTitle(g_MainHost)
    RedrawAnyWindow(g_MainHost.hwnd)
}

; Switches to newly stacked tab after delay; lets content load before showing to reduce glitch.
SwitchToNewTabDelayed(host, tabId, *) {
    if host.HasProp("pendingSwitchTimer")
        host.pendingSwitchTimer := ""
    if !host || !host.tabRecords.Has(tabId)
        return
    host.activeTabId := tabId
    ShowOnlyActiveTab(host)
    UpdateHostTitle(host)
    if host.tabRecords.Has(tabId) && IsWindowExists(host.tabRecords[tabId].contentHwnd)
        FocusEmbeddedContent(host.hwnd, host.tabRecords[tabId].contentHwnd)
}

; Shell hook handler: wParam 1=created, 2=destroyed; lParam=hwnd.
; HSHELL_REDRAW (6) not used: it fires for taskbar windows; embedded windows are reparented
; and typically no longer in the taskbar, so we wouldn't receive it. RefreshWindows already
; updates tab titles periodically via GetPreferredTabTitle.
OnShellHook(wParam, lParam, msg, hwnd) {
    if (wParam = 1)
        OnWindowCreated(lParam)
    else if (wParam = 2)
        OnWindowDestroyed(lParam)
}

; WinEvent hook callback. Three narrow hooks are registered, one per event:
;   0x8002 EVENT_OBJECT_SHOW       — window becomes visible (fires before taskbar registration)
;   0x800C EVENT_OBJECT_NAMECHANGE — title set/changed (WPF often sets title after show)
;   0x8018 EVENT_OBJECT_UNCLOAKED  — UWP/WinUI window uncloaks
; idObject=0 (OBJID_WINDOW) means the event is for the window object itself, not a child control.
WinEventProc(hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime) {
    if (event != 0x8002 && event != 0x800C && event != 0x8018)
        return
    if (idObject != 0 || !hwnd)
        return
    ; NAMECHANGE on already-tracked tab: update tab title and switch to that tab
    if (event = 0x800C) {
        for host in GetAllHosts() {
            for tabId, record in host.tabRecords {
                if (record.topHwnd = hwnd || record.contentHwnd = hwnd) {
                    newTitle := GetPreferredTabTitle(record)
                    if newTitle != "" && newTitle != record.title {
                        record.title := newTitle
                        LayoutTabButtons(host)
                        ; SelectTab → ShowOnlyActiveTab → UpdateTabButtonStyles → WinRedraw handles
                        ; the visual refresh; a separate RedrawHostWindow here would cause a double repaint.
                        SelectTab(host, tabId)
                    }
                    return
                }
            }
        }
    }
    OnWindowCreated(hwnd)
}

; Shell hook: when a new window is created, try to add it as a tab if it matches.
OnWindowCreated(hwnd) {
    global g_MainHost, g_PendingCandidates, g_IsCleaningUp
    if !g_MainHost || g_IsCleaningUp
        return
    ; Fast HWND check before the expensive candidate build (which walks all descendants).
    ; Covers NAMECHANGE events on already-tracked tabs whose title changed.
    for host in GetAllHosts() {
        for tabId, record in host.tabRecords {
            if (record.topHwnd = hwnd || record.contentHwnd = hwnd)
                return
        }
    }
    for tabId, pending in g_PendingCandidates {
        if pending.candidate.topHwnd = hwnd
            return
    }
    candidate := BuildCandidateFromTopWindow(hwnd)
    if !IsObject(candidate)
        return
    ; Skip if already embedded or pending
    for host in GetAllHosts() {
        if host.tabRecords.Has(candidate.id)
            return
    }
    if g_PendingCandidates.Has(candidate.id)
        return
    ; Skip if content/top already embedded
    for host in GetAllHosts() {
        for tabId, record in host.tabRecords {
            if (record.contentHwnd = candidate.contentHwnd || record.topHwnd = candidate.topHwnd)
                return
        }
    }
    TryStackOrPending(g_MainHost, candidate)
}

; Shell hook: removes tab from pending or host when its window is destroyed.
OnWindowDestroyed(hwnd) {
    global g_MainHost, g_PendingCandidates, g_ShowOnlyWhenTabs
    ; Remove from pending if this hwnd was a candidate
    stalePending := []
    for tabId, pending in g_PendingCandidates {
        if (pending.candidate.topHwnd = hwnd)
            stalePending.Push(tabId)
    }
    for tabId in stalePending
        g_PendingCandidates.Delete(tabId)
    ; Find and remove any tab that had this window
    for host in GetAllHosts() {
        for tabId, record in host.tabRecords {
            if (record.topHwnd = hwnd || record.contentHwnd = hwnd) {
                ; Window still alive = reparented/hidden by StackTabs, not closed by user.
                ; The slow sweep handles truly destroyed windows via stale cleanup.
                if IsWindowExists(hwnd)
                    return
                RemoveTrackedTab(host, tabId, false)
                ; Update layout and visibility
                LayoutTabButtons(host)
                ShowOnlyActiveTab(host)
                UpdateHostTitle(host)
                RedrawAnyWindow(host.hwnd)
                if g_MainHost && host = g_MainHost && g_ShowOnlyWhenTabs && host.tabOrder.Length = 0
                    g_MainHost.gui.Hide()
                return
            }
        }
    }
}

; Slow-sweep timer: discovers new windows, updates titles, removes stale tabs, shows/hides host.
RefreshWindows(*) {
    global g_MainHost, g_IsCleaningUp, g_PendingCandidates, g_TabDisappearGraceMs, g_ShowOnlyWhenTabs

    if !g_MainHost || g_IsCleaningUp
        return

    now := A_TickCount

    ; Update all hosts: keep tabs alive, check for stale tabs
    for host in GetAllHosts() {
        ; Use IsWindow: WinExist returns 0 for hidden windows; host may be hidden when ShowOnlyWhenTabs
        if !IsWindowExists(host.hwnd)
            continue

        structureChanged := false
        titleChanged := false
        tabIdOfLastTitleChange := ""
        currentIds := Map()

        ; Keep existing embedded tabs alive
        for tabId in host.tabOrder {
            if !host.tabRecords.Has(tabId)
                continue
            record := host.tabRecords[tabId]
            if IsWindowExists(record.contentHwnd) {
                record.lastSeenTick := now
                title := GetPreferredTabTitle(record)
                if title != "" && title != record.title {
                    record.title := title
                    titleChanged := true
                    tabIdOfLastTitleChange := tabId
                }
            }
        }

        ; Discovery only for main host (popouts don't scan for new windows)
        if host = g_MainHost {
            candidates := DiscoverCandidateWindows()
            for candidate in candidates {
                currentIds[candidate.id] := true

                if host.tabRecords.Has(candidate.id) {
                    if UpdateTrackedTab(host, candidate.id, candidate)
                        structureChanged := true
                    continue
                }

                if !g_PendingCandidates.Has(candidate.id) {
                    TryStackOrPending(host, candidate)
                    if host.tabRecords.Has(candidate.id)
                        structureChanged := true
                    continue
                }

                ; Already pending: refresh candidate metadata only, do NOT reset firstSeen
                g_PendingCandidates[candidate.id].candidate := candidate
            }

            stalePending := []
            for tabId, pending in g_PendingCandidates {
                if !currentIds.Has(tabId)
                    stalePending.Push(tabId)
            }
            for tabId in stalePending {
                g_PendingCandidates.Delete(tabId)
            }
        }

        ; Stale tab cleanup for this host
        staleTabs := []
        for tabId in host.tabOrder {
            if !host.tabRecords.Has(tabId)
                continue
            record := host.tabRecords[tabId]
            winExists := IsWindowExists(record.contentHwnd)
            if winExists
                continue
            if (now - record.lastSeenTick) > g_TabDisappearGraceMs
                staleTabs.Push(tabId)
        }

        for tabId in staleTabs {
            RemoveTrackedTab(host, tabId, false)
            structureChanged := true
        }

        if host.activeTabId && !host.tabRecords.Has(host.activeTabId)
            host.activeTabId := ""
        if (host.activeTabId = "") && host.tabOrder.Length
            host.activeTabId := host.tabOrder[1]

        if structureChanged || titleChanged
            LayoutTabButtons(host)
        if structureChanged
            RedrawAnyWindow(host.hwnd)
        ; Title changed: redraw tab bar to show the new label, but do NOT switch tabs.
        ; Switching focus away from the user's current tab on a background title change is disruptive.
        if titleChanged && !structureChanged
            DrawTabBar(host)

        ; Only refresh content/tabs when something changed to avoid flickering
        needsContentRefresh := structureChanged || (host.activeTabId != (host.HasProp("lastRefreshActiveTabId") ? host.lastRefreshActiveTabId : ""))
        if needsContentRefresh {
            host.lastRefreshActiveTabId := host.activeTabId
            ShowOnlyActiveTab(host)
        }
        UpdateHostTitle(host)
    }

    ; When ShowOnlyWhenTabs: hide when 0 tabs, show when 1+ tabs (only if hidden).
    if g_ShowOnlyWhenTabs && g_MainHost && IsWindowExists(g_MainHost.hwnd) {
        liveCount := 0
        for tabId in g_MainHost.tabOrder {
            if g_MainHost.tabRecords.Has(tabId) && IsWindowExists(g_MainHost.tabRecords[tabId].contentHwnd)
                liveCount++
        }
        if liveCount >= 1 {
            if !WinExist("ahk_id " g_MainHost.hwnd)
                g_MainHost.gui.Show("NoActivate")
        } else
            g_MainHost.gui.Hide()
    }
}

; Scans all top-level windows and returns those matching title patterns and not already embedded.
DiscoverCandidateWindows() {
    candidates := []
    seenIds := Map()

    ; Build set of all embedded HWNDs (content + top) across all hosts
    embeddedHwnds := Map()
    for host in GetAllHosts() {
        if host.hwnd
            embeddedHwnds[host.hwnd ""] := true
        for tabId, record in host.tabRecords {
            embeddedHwnds[record.contentHwnd ""] := true
            if record.topHwnd != record.contentHwnd
                embeddedHwnds[record.topHwnd ""] := true
        }
    }

    for hwnd in WinGetList() {
        if embeddedHwnds.Has(hwnd "")
            continue

        candidate := BuildCandidateFromTopWindow(hwnd)
        if !IsObject(candidate)
            continue
        ; Skip if this candidate's content/top is already embedded anywhere
        if embeddedHwnds.Has(candidate.contentHwnd "") || embeddedHwnds.Has(candidate.topHwnd "")
            continue
        if seenIds.Has(candidate.id)
            continue

        seenIds[candidate.id] := true
        candidates.Push(candidate)
    }

    return candidates
}

; Builds a candidate object from a top-level hwnd if it matches title patterns and size.
BuildCandidateFromTopWindow(topHwnd) {
    global g_WindowTitleMatches, g_TargetExe, g_DebugDiscovery

    try {
        if !IsWindowExists(topHwnd)
            return ""
        title := WinGetTitle("ahk_id " topHwnd)
        if !DllCall("IsWindowVisible", "ptr", topHwnd)
            return ""
        if (title = "")
            return ""
        if g_WindowTitleMatches.Length = 0
            return ""  ; No match patterns configured
        matched := false
        for pat in g_WindowTitleMatches {
            if InStr(title, pat, false) {
                matched := true
                break
            }
        }
        if !matched
            return ""

        processName := WinGetProcessName("ahk_id " topHwnd)
        ; Exclude processes that crash or misbehave when reparented (e.g. explorer.exe)
        if StrLower(processName) = "explorer.exe"
            return ""
        if g_TargetExe && (StrLower(processName) != StrLower(g_TargetExe))
            return ""

        WinGetPos(, , &w, &h, "ahk_id " topHwnd)
        if (w < 120 || h < 80)
            return ""

        contentHwnd := FindStableContentWindow(topHwnd)
        if !contentHwnd
            contentHwnd := topHwnd

        candidate := {
            id: BuildCandidateId(topHwnd, title, processName, contentHwnd),
            title: title,
            topHwnd: topHwnd,
            contentHwnd: contentHwnd,
            processName: processName,
            rootOwner: GetRootOwner(topHwnd),
            hierarchySummary: g_DebugDiscovery ? DescribeWindowHierarchy(topHwnd, contentHwnd) : ""
        }
        return candidate
    } catch {
        return ""
    }
}

; Picks the best child window to embed (largest, title-matching, not dialog/static).
FindStableContentWindow(topHwnd) {
    bestHwnd := topHwnd
    bestScore := ScoreContentCandidate(topHwnd, topHwnd)

    for childHwnd in GetDescendantWindows(topHwnd) {
        score := ScoreContentCandidate(topHwnd, childHwnd)
        if score > bestScore {
            bestScore := score
            bestHwnd := childHwnd
        }
    }

    return bestHwnd
}

; Scores a window as content candidate: area + title match bonus, minus dialog/static penalty.
ScoreContentCandidate(topHwnd, hwnd) {
    global g_WindowTitleMatches

    if !IsWindowExists(hwnd)
        return -1
    if !DllCall("IsWindowVisible", "ptr", hwnd)
        return -1

    try WinGetPos(, , &w, &h, "ahk_id " hwnd)
    catch
        return -1

    if (w < 80 || h < 40)
        return -1

    area := w * h
    title := SafeWinGetTitle(hwnd)
    className := GetWindowClassName(hwnd)
    score := area

    if hwnd != topHwnd
        score += 1000000
    titleMatches := false
    for pat in g_WindowTitleMatches {
        if (title != "") && InStr(title, pat, false) {
            titleMatches := true
            break
        }
    }
    if titleMatches
        score += 500000
    if className = "#32770"
        score -= 250000
    if (className = "Static" || className = "Button")
        score -= 900000

    return score
}

; Returns all descendant windows of a parent (recursive).
GetDescendantWindows(parentHwnd) {
    result := []
    visited := Map()
    CollectDescendantWindows(parentHwnd, &result, visited)
    return result
}

; Recursively collects child hwnds into result, avoiding cycles.
CollectDescendantWindows(parentHwnd, &result, visited) {
    try childWindows := WinGetControlsHwnd("ahk_id " parentHwnd)
    catch
        return

    for childHwnd in childWindows {
        key := childHwnd ""
        if visited.Has(key)
            continue

        visited[key] := true
        result.Push(childHwnd)
        CollectDescendantWindows(childHwnd, &result, visited)
    }
}

; Builds unique tab ID from process, owner, title, content class, and content hwnd.
BuildCandidateId(topHwnd, title, processName, contentHwnd) {
    rootOwner := GetRootOwner(topHwnd)
    contentClass := GetWindowClassName(contentHwnd)
    ; Include contentHwnd so multiple windows with same title (e.g. 3x PowerShell) get unique IDs
    return StrLower(processName) "|" rootOwner "|" NormalizeTitle(title) "|" contentClass "|" contentHwnd
}

; Shows duplicate-window dialog; returns "Yes" or "No". Y/N keys work regardless of locale.
DuplicateConfirmDialog(shortTitle) {
    result := ""
    dlg := Gui("+AlwaysOnTop +ToolWindow", "StackTabs")
    dlg.Add("Text", "w350", "We detected a duplicate window.`n`n" shortTitle "`n`nWant to close the older one?")
    btnY := dlg.Add("Button", "Default w90 h28", "Yes (Y)")
    btnN := dlg.Add("Button", "w90 h28 x+10", "No (N)")

    cleanup(*) {
        try Hotkey("y", "Off")
        try Hotkey("Y", "Off")
        try Hotkey("n", "Off")
        try Hotkey("N", "Off")
    }

    submitYes(*) {
        if result = ""
            result := "Yes"
        cleanup()
        try dlg.Destroy()
    }
    submitNo(*) {
        if result = ""
            result := "No"
        cleanup()
        try dlg.Destroy()
    }

    dlg.OnEvent("Close", (*) => (result := "No", cleanup(), dlg.Destroy()))
    btnY.OnEvent("Click", submitYes)
    btnN.OnEvent("Click", submitNo)

    dlg.Show()
    dlgHwnd := dlg.Hwnd

    fnY(*) {
        if dlgHwnd && IsWindowExists(dlgHwnd) && WinActive("ahk_id " dlgHwnd)
            submitYes()
    }
    fnN(*) {
        if dlgHwnd && IsWindowExists(dlgHwnd) && WinActive("ahk_id " dlgHwnd)
            submitNo()
    }

    Hotkey("y", fnY, "On")
    Hotkey("Y", fnY, "On")
    Hotkey("n", fnN, "On")
    Hotkey("N", fnN, "On")

    try {
        while IsWindowExists(dlgHwnd)
            Sleep(50)
    } finally {
        cleanup()
    }
    return result = "" ? "No" : result
}

; Returns tabId of oldest tab with same process and normalized title, or "" if none.
FindTabWithSameTitle(host, candidate) {
    newNorm := NormalizeTitle(candidate.title)
    newProc := StrLower(candidate.processName)
    for tabId in host.tabOrder {
        if !host.tabRecords.Has(tabId)
            continue
        record := host.tabRecords[tabId]
        if (StrLower(record.processName) = newProc && NormalizeTitle(record.title) = newNorm)
            return tabId
    }
    return ""
}

; Adds candidate as a tracked tab: builds record, attaches window, updates layout.
CreateTrackedTab(host, candidate) {
    if host.tabRecords.Has(candidate.id)
        return false

    record := BuildTrackedRecord(candidate)
    host.tabRecords[candidate.id] := record

    if !AttachTrackedWindow(host, candidate.id) {
        host.tabRecords.Delete(candidate.id)
        return false
    }

    host.tabOrder.Push(candidate.id)
    if (host.activeTabId = "")
        host.activeTabId := candidate.id
    return true
}

; Builds a tab record with original position, style, owner for later restore.
BuildTrackedRecord(candidate) {
    WinGetPos(&x, &y, &w, &h, "ahk_id " candidate.contentHwnd)
    parentHwnd := DllCall("GetParent", "ptr", candidate.contentHwnd, "ptr")
    ; For child windows: SetWindowPos expects parent-relative coords. Store both for fallback.
    if parentHwnd {
        point := Buffer(8)
        NumPut("int", x, point, 0)
        NumPut("int", y, point, 4)
        DllCall("MapWindowPoints", "ptr", 0, "ptr", parentHwnd, "ptr", point, "uint", 1)
        restoreX := NumGet(point, 0, "int")
        restoreY := NumGet(point, 4, "int")
    } else {
        restoreX := x
        restoreY := y
    }

    return {
        id: candidate.id,
        title: candidate.title,
        topHwnd: candidate.topHwnd,
        contentHwnd: candidate.contentHwnd,
        processName: candidate.processName,
        rootOwner: candidate.rootOwner,
        originalContentParent: parentHwnd,
        originalContentOwner: GetWindowLongPtrValue(candidate.contentHwnd, -8),
        originalContentStyle: GetWindowLongPtrValue(candidate.contentHwnd, -16),
        originalContentExStyle: GetWindowLongPtrValue(candidate.contentHwnd, -20),
        originalContentX: restoreX,
        originalContentY: restoreY,
        originalContentScreenX: x,
        originalContentScreenY: y,
        originalContentW: w,
        originalContentH: h,
        sourceWasHidden: false,
        sourceWasVisible: (candidate.topHwnd != candidate.contentHwnd) && DllCall("IsWindowVisible", "ptr", candidate.topHwnd) ? 1 : 0,
        lastSeenTick: A_TickCount
    }
}

; Updates tab record and re-attaches if top/content changed or window was reparented.
UpdateTrackedTab(host, tabId, candidate) {
    record := host.tabRecords[tabId]
    record.lastSeenTick := A_TickCount
    record.title := candidate.title
    record.processName := candidate.processName
    record.rootOwner := candidate.rootOwner

    if (record.topHwnd != candidate.topHwnd || record.contentHwnd != candidate.contentHwnd) {
        RebindTrackedTab(host, tabId, candidate)
        return true
    }

    if IsWindowExists(record.contentHwnd) {
        currentParent := DllCall("GetParent", "ptr", record.contentHwnd, "ptr")
        if currentParent != host.clientHwnd {
            AttachTrackedWindow(host, tabId)
            return true
        }
    }

    return false
}

; Detaches and re-attaches tab with new candidate (top/content changed).
RebindTrackedTab(host, tabId, candidate) {
    if !host.tabRecords.Has(tabId)
        return

    DetachTrackedWindow(host, tabId, false, false)

    record := BuildTrackedRecord(candidate)
    host.tabRecords[tabId] := record
    AttachTrackedWindow(host, tabId)
    AppendDebugLog("Rebound tab: " tabId "`r`n" candidate.hierarchySummary "`r`n")
}

; Closes the window and removes the tab from the host.
CloseTab(host, tabId) {
    if !host.tabRecords.Has(tabId)
        return

    record := host.tabRecords[tabId]
    topHwnd := record.topHwnd
    contentHwnd := record.contentHwnd
    ; Close before detach - window may process close better while still embedded
    CloseWindowReliably(topHwnd, contentHwnd)
    RemoveTrackedTab(host, tabId, false)

    SetTimer(CloseTabDeferredUpdate.Bind(host), -1)
}

; Deferred layout after close; destroys empty popout host.
CloseTabDeferredUpdate(host, *) {
    global g_PopoutHosts
    if !host.hwnd || !IsWindowExists(host.hwnd)
        return
    LayoutTabButtons(host)
    ShowOnlyActiveTab(host)
    UpdateHostTitle(host)
    RedrawAnyWindow(host.hwnd)
    if host.hwnd && IsWindowExists(host.hwnd) && host.tabOrder.Length > 0
        WinActivate("ahk_id " host.hwnd)
    ; Destroy empty popout
    if host.isPopout && host.tabOrder.Length = 0 {
        for i, h in g_PopoutHosts {
            if h = host {
                g_PopoutHosts.RemoveAt(i)
                break
            }
        }
        InvalidateHostsCache()
        host.gui.Destroy()
    }
}

; Closes the currently active tab.
CloseActiveTab(host) {
    if host.activeTabId != ""
        CloseTab(host, host.activeTabId)
}

; Detaches tab, removes from order/records; optionally restores window to original position.
RemoveTrackedTab(host, tabId, restoreWindow := true) {
    global g_ShowOnlyWhenTabs
    if !host.tabRecords.Has(tabId)
        return

    DetachTrackedWindow(host, tabId, restoreWindow, true)

    closedIdx := 0
    for idx, currentId in host.tabOrder {
        if currentId = tabId {
            closedIdx := idx
            host.tabOrder.RemoveAt(idx)
            break
        }
    }

    if host.tabRecords.Has(tabId)
        host.tabRecords.Delete(tabId)

    ; Focus the tab to the left of the closed one; if closed was first, focus the new first.
    if host.activeTabId = tabId && host.tabOrder.Length
        host.activeTabId := host.tabOrder[Max(1, closedIdx - 1)]
    else if host.activeTabId = tabId
        host.activeTabId := ""

    ; When ShowOnlyWhenTabs and main host now has 0 tabs, hide immediately
    if !host.isPopout && g_ShowOnlyWhenTabs && host.tabOrder.Length = 0
        host.gui.Hide()
}

; Reparents content window into host client area; hides top window if different.
AttachTrackedWindow(host, tabId) {
    if !host.tabRecords.Has(tabId)
        return false

    record := host.tabRecords[tabId]
    hwnd := record.contentHwnd

    if !IsWindowExists(hwnd)
        return false

    ; Grab focus before hiding the original window so Windows doesn't redirect it elsewhere.
    if host.hwnd && IsWindowExists(host.hwnd)
        DllCall("SetForegroundWindow", "ptr", host.hwnd)

    try SendMessage(0x000B, 0, 0,, "ahk_id " hwnd,,,, 500)  ; WM_SETREDRAW FALSE

    if (record.topHwnd != hwnd) && IsWindowExists(record.topHwnd) {
        record.sourceWasHidden := true
        DllCall("ShowWindow", "ptr", record.topHwnd, "int", SW_HIDE)
    }

    newStyle := record.originalContentStyle
    newStyle |= 0x40000000 ; WS_CHILD
    newStyle &= ~0x80000000 ; WS_POPUP
    newStyle &= ~0x00C00000 ; WS_CAPTION
    newStyle &= ~0x00040000 ; WS_THICKFRAME
    newStyle &= ~0x00020000 ; WS_MINIMIZEBOX
    newStyle &= ~0x00010000 ; WS_MAXIMIZEBOX

    newExStyle := record.originalContentExStyle
    newExStyle &= ~0x00040000 ; WS_EX_APPWINDOW
    newExStyle &= ~0x00000200 ; WS_EX_CLIENTEDGE
    newExStyle &= ~0x00000100 ; WS_EX_WINDOWEDGE

    prevOwner := SetWindowLongPtrValue(hwnd, -8, 0)
    prevStyle  := SetWindowLongPtrValue(hwnd, -16, newStyle)
    prevExStyle := SetWindowLongPtrValue(hwnd, -20, newExStyle)
    newParent := DllCall("SetParent", "ptr", hwnd, "ptr", host.clientHwnd, "ptr")
    if !newParent {
        ; Reparent failed — undo style changes to leave window in its original state
        SetWindowLongPtrValue(hwnd, -8, prevOwner)
        SetWindowLongPtrValue(hwnd, -16, record.originalContentStyle)
        SetWindowLongPtrValue(hwnd, -20, record.originalContentExStyle)
        AppendDebugLog("AttachTrackedWindow: SetParent failed for hwnd=" hwnd " tabId=" tabId "`r`n", true)
        SendMessage(0x000B, 1, 0,, "ahk_id " hwnd,,,, 500)  ; WM_SETREDRAW TRUE (re-enable drawing)
        return false
    }

    ; Position at final content rect immediately so when we show it there's no resize glitch.
    GetEmbedRect(host, &areaX, &areaY, &areaW, &areaH)
    areaX += 1
    areaY += 1
    areaW -= 2
    areaH -= 2
    flags := SWP_FRAMECHANGED | SWP_SHOWWINDOW | SWP_NOZORDER | SWP_NOACTIVATE
    DllCall("SetWindowPos", "ptr", hwnd, "ptr", 0, "int", areaX, "int", areaY, "int", areaW, "int", areaH, "uint", flags)
    try SendMessage(0x000B, 1, 0,, "ahk_id " hwnd,,,, 500)  ; WM_SETREDRAW TRUE
    DllCall("ShowWindow", "ptr", hwnd, "int", SW_HIDE)
    return true
}

; Reparents content back to original parent, restores style/position; shows top if it was hidden.
DetachTrackedWindow(host, tabId, restoreWindow := true, restoreSource := true) {
    if !host.tabRecords.Has(tabId)
        return

    record := host.tabRecords[tabId]
    hwnd := record.contentHwnd

    ; When closing, claim host focus before any ShowWindow so focus stays in StackTabs.
    if !restoreWindow && host.hwnd && IsWindowExists(host.hwnd)
        DllCall("SetForegroundWindow", "ptr", host.hwnd)

    ; Show parent FIRST (critical for WinUI/XAML apps like PowerShell/Windows Terminal)
    ; so the composition tree can reattach before we reparent the content.
    ; Skip when closing: topHwnd is about to close anyway and showing it steals focus.
    if restoreSource && restoreWindow && record.sourceWasHidden && (record.topHwnd != hwnd) && IsWindowExists(record.topHwnd) {
        DllCall("ShowWindow", "ptr", record.topHwnd, "int", record.sourceWasVisible ? SW_SHOWNOACTIVATE : SW_HIDE)
    }

    if IsWindowExists(hwnd) {
        ; Validate parent: if destroyed, fall back to desktop (top-level window)
        parentHwnd := record.originalContentParent
        if !parentHwnd || !IsWindowExists(parentHwnd)
            parentHwnd := 0

        ; Parent valid: use parent-relative coords. Parent 0 (fallback): use screen coords.
        if parentHwnd {
            posX := record.originalContentX
            posY := record.originalContentY
        } else if record.HasProp("originalContentScreenX") {
            posX := record.originalContentScreenX
            posY := record.originalContentScreenY
        } else {
            posX := record.originalContentX
            posY := record.originalContentY
        }

        newParent := DllCall("SetParent", "ptr", hwnd, "ptr", parentHwnd, "ptr")
        if !newParent
            AppendDebugLog("DetachTrackedWindow: SetParent failed for hwnd=" hwnd " tabId=" tabId "`r`n", true)
        SetWindowLongPtrValue(hwnd, -8, record.originalContentOwner)
        SetWindowLongPtrValue(hwnd, -16, record.originalContentStyle)
        SetWindowLongPtrValue(hwnd, -20, record.originalContentExStyle)

        flags := SWP_FRAMECHANGED | SWP_SHOWWINDOW
        DllCall("SetWindowPos", "ptr", hwnd, "ptr", 0
            , "int", posX, "int", posY
            , "int", record.originalContentW, "int", record.originalContentH
            , "uint", flags)
        DllCall("ShowWindow", "ptr", hwnd, "int", restoreWindow ? SW_SHOW : SW_HIDE)
    }
    record.sourceWasHidden := false
}

; Moves embedded window from source host to dest host (popout/merge).
TransferTrackedWindow(sourceHost, destHost, tabId) {
    if !destHost.tabRecords.Has(tabId)
        return false

    record := destHost.tabRecords[tabId]
    hwnd := record.contentHwnd

    if !IsWindowExists(hwnd)
        return false

    ; Direct reparent: source host's client -> dest host's client
    ; Window stays as WS_CHILD the whole time - no restore to original
    newParent := DllCall("SetParent", "ptr", hwnd, "ptr", destHost.clientHwnd, "ptr")
    if !newParent {
        AppendDebugLog("TransferTrackedWindow: SetParent failed for hwnd=" hwnd " tabId=" tabId "`r`n", true)
        return false
    }

    flags := SWP_FRAMECHANGED | SWP_SHOWWINDOW | SWP_NOZORDER | SWP_NOACTIVATE
    DllCall("SetWindowPos", "ptr", hwnd, "ptr", 0, "int", 0, "int", 0, "int", 100, "int", 100, "uint", flags)
    DllCall("ShowWindow", "ptr", hwnd, "int", SW_SHOWNOACTIVATE)
    RedrawAnyWindow(hwnd)
    return true
}

; Positions and sizes tab buttons, popout/close controls, and indicators.
LayoutTabButtons(host, windowWidth := 0, windowHeight := 0) {
    global g_HostWidth, g_HostHeight, g_HostPadding, g_TabGap, g_MinTabWidth, g_MaxTabWidth, g_TabHeight
    global g_CloseButtonWidth, g_PopoutButtonWidth, g_HeaderHeight, g_TabBarAlignment, g_TabBarOffsetY
    global g_UseCustomTitleBar, g_TitleBarHeight, g_TabPosition

    if !host || !host.gui
        return
    if !host.hwnd || !IsWindowExists(host.hwnd)
        return

    host.isLayingOut := true
    if host.hwnd && IsWindowExists(host.hwnd) {
        prev := DetectHiddenWindows(true)
        try SendMessage(0x000B, 0, 0,, "ahk_id " host.hwnd)
        DetectHiddenWindows(prev)
    }
    try {
    if !windowWidth {
        windowWidth := GetClientWidth(host.hwnd)
        if !windowWidth
            windowWidth := g_HostWidth
    }
    if !windowWidth
        windowWidth := g_HostWidth

    if g_TabPosition = "bottom" {
        if !windowHeight {
            windowHeight := GetClientHeight(host.hwnd)
            if !windowHeight
                windowHeight := g_HostHeight
        }
        tabBarY := windowHeight - g_HeaderHeight
    } else {
        tabBarY := g_UseCustomTitleBar ? g_TitleBarHeight : 0
    }
    if host.HasProp("titleBarBg") && host.titleBarBg {
        host.titleBarBg.Move(0, 0, windowWidth, g_TitleBarHeight)
        host.titleText.Move(8, 0, windowWidth - 60, g_TitleBarHeight)
        host.titleCloseBtn.Move(windowWidth - 46, 0, 46, g_TitleBarHeight)
    }
    ; Resize tab bar background to full width
    if host.HasProp("tabBarBg") && host.tabBarBg
        host.tabBarBg.Move(0, tabBarY, windowWidth, g_HeaderHeight)

    tabCount := host.tabOrder.Length

    if !tabCount {
        DrawTabBar(host)
        return
    }

    global g_ThemeFontName, g_ThemeFontNameTab, g_ThemeFontSize, g_ThemeIconColor
    global g_ThemeIconFont, g_ThemeIconFontSize, g_IconClose, g_IconPopout, g_IconMerge
    ; Compute tab Y: legacy TabBarOffsetY >= 0, or from alignment (top/center/bottom)
    if g_TabBarOffsetY >= 0
        tabOffsetY := g_TabBarOffsetY
    else {
        align := StrLower(g_TabBarAlignment)
        if (align = "top")
            tabOffsetY := 0
        else if (align = "bottom")
            tabOffsetY := g_HeaderHeight - g_TabHeight
        else
            tabOffsetY := (g_HeaderHeight - g_TabHeight) // 2  ; center (default)
    }
    DrawTabBar(host)
    } finally {
    if host.hwnd && IsWindowExists(host.hwnd) {
        prev := DetectHiddenWindows(true)
        try SendMessage(0x000B, 1, 0,, "ahk_id " host.hwnd)
        DetectHiddenWindows(prev)
        DllCall("RedrawWindow", "Ptr", host.hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x0085)
    }
        host.isLayingOut := false
    }
}

; Sets active tab, shows its content, updates host title.
SelectTab(host, tabId, *) {
    if !host.tabRecords.Has(tabId)
        return

    host.activeTabId := tabId
    ShowOnlyActiveTab(host)
    UpdateHostTitle(host)
    if host.tabRecords.Has(tabId) && IsWindowExists(host.tabRecords[tabId].contentHwnd)
        FocusEmbeddedContent(host.hwnd, host.tabRecords[tabId].contentHwnd)
}

; Moves tab to a new popout host window; positions it side-by-side with source.
PopOutTab(sourceHost, tabId) {
    global g_MainHost, g_PopoutHosts

    if !sourceHost.tabRecords.Has(tabId)
        return

    record := sourceHost.tabRecords[tabId]

    ; Create pop-out host and move record
    popoutHost := BuildHostInstance(true)
    popoutHost.tabRecords[tabId] := record
    popoutHost.tabOrder.Push(tabId)
    popoutHost.activeTabId := tabId

    ; Remove from source
    closedIdx := 0
    for idx, currentId in sourceHost.tabOrder {
        if currentId = tabId {
            closedIdx := idx
            sourceHost.tabOrder.RemoveAt(idx)
            break
        }
    }
    sourceHost.tabRecords.Delete(tabId)
    if sourceHost.activeTabId = tabId && sourceHost.tabOrder.Length
        sourceHost.activeTabId := sourceHost.tabOrder[Max(1, closedIdx - 1)]
    else if sourceHost.activeTabId = tabId
        sourceHost.activeTabId := ""

    if !TransferTrackedWindow(sourceHost, popoutHost, tabId) {
        ; Failed - window never moved, restore data structures
        sourceHost.tabRecords[tabId] := record
        sourceHost.tabOrder.Push(tabId)
        sourceHost.activeTabId := tabId
        g_PopoutHosts.Pop()
        InvalidateHostsCache()
        popoutHost.gui.Destroy()
        return
    }

    ; Position pop-out host side-by-side with source
    ArrangeHostsSideBySide(sourceHost, popoutHost)

    ; Defer layout to next message pump so we don't move controls under the cursor
    ; (which can cause a spurious click on the remaining tab's popout button)
    SetTimer(PopOutTabDeferredLayout.Bind(sourceHost, popoutHost), -1)
}

; Deferred layout after popout: refresh both hosts to avoid click-through issues.
PopOutTabDeferredLayout(sourceHost, popoutHost, *) {
    LayoutTabButtons(sourceHost)
    ShowOnlyActiveTab(sourceHost)
    UpdateHostTitle(sourceHost)
    LayoutTabButtons(popoutHost)
    ShowOnlyActiveTab(popoutHost)
    UpdateHostTitle(popoutHost)
    RedrawAnyWindow(sourceHost.hwnd)
    RedrawAnyWindow(popoutHost.hwnd)
}

MergeBackTab(popoutHost, tabId) {
    global g_MainHost, g_PopoutHosts

    if !popoutHost.tabRecords.Has(tabId) || !popoutHost.isPopout
        return

    record := popoutHost.tabRecords[tabId]

    ; Add to main host
    g_MainHost.tabRecords[tabId] := record
    g_MainHost.tabOrder.Push(tabId)
    if (g_MainHost.activeTabId = "")
        g_MainHost.activeTabId := tabId

    ; Remove from popout
    for idx, currentId in popoutHost.tabOrder {
        if currentId = tabId {
            popoutHost.tabOrder.RemoveAt(idx)
            break
        }
    }
    popoutHost.tabRecords.Delete(tabId)

    if !TransferTrackedWindow(popoutHost, g_MainHost, tabId) {
        ; Failed - window never moved, restore data structures
        g_MainHost.tabRecords.Delete(tabId)
        for idx, currentId in g_MainHost.tabOrder {
            if currentId = tabId {
                g_MainHost.tabOrder.RemoveAt(idx)
                break
            }
        }
        if (g_MainHost.activeTabId = tabId)
            g_MainHost.activeTabId := g_MainHost.tabOrder.Length ? g_MainHost.tabOrder[1] : ""
        popoutHost.tabRecords[tabId] := record
        popoutHost.tabOrder.Push(tabId)
        popoutHost.activeTabId := tabId
        return
    }

    ; Destroy popout host
    for i, h in g_PopoutHosts {
        if h = popoutHost {
            g_PopoutHosts.RemoveAt(i)
            break
        }
    }
    InvalidateHostsCache()
    popoutHost.gui.Destroy()

    LayoutTabButtons(g_MainHost)
    ShowOnlyActiveTab(g_MainHost)
    UpdateHostTitle(g_MainHost)
    RedrawAnyWindow(g_MainHost.hwnd)
}

; Positions host2 (popout) on the opposite half of the monitor from host1.
ArrangeHostsSideBySide(host1, host2) {
    try {
        WinGetPos(&x1, &y1, &w1, &h1, "ahk_id " host1.hwnd)

        ; Work area of the monitor containing host1
        cx := x1 + w1 // 2
        cy := y1 + h1 // 2
        workL := 0, workT := 0, workR := A_ScreenWidth, workB := A_ScreenHeight
        Loop MonitorGetCount() {
            MonitorGetWorkArea(A_Index, &mL, &mT, &mR, &mB)
            if (cx >= mL && cx < mR && cy >= mT && cy < mB) {
                workL := mL, workT := mT, workR := mR, workB := mB
                break
            }
        }

        ; Place popout in the opposite half when main host is tiled (avoids off-screen or cramped placement)
        workW := workR - workL
        hostCenter := x1 + w1 // 2
        workCenter := workL + workW // 2

        if (hostCenter >= workCenter) {
            ; Main host is on right half â€” put popout on left half
            x2 := workL
            y2 := Max(workT, Min(y1, workB - h1))
        } else {
            ; Main host is on left half â€” put popout on right half
            x2 := workR - w1
            y2 := Max(workT, Min(y1, workB - h1))
        }
        x2 := Max(workL, Min(x2, workR - w1))
        y2 := Max(workT, Min(y2, workB - h1))

        host2.gui.Show("x" x2 " y" y2 " w" w1 " h" h1)
    } catch {
        host2.gui.Show()
    }
}

; Shows active tab content, hides others; positions content area and border.
ShowOnlyActiveTab(host) {
    global g_ThemeContentBorder
    DllCall("LockWindowUpdate", "Ptr", host.hwnd)
    try {
    if (host.activeTabId != "") && !host.tabRecords.Has(host.activeTabId)
        host.activeTabId := ""
    if (host.activeTabId = "") && host.tabOrder.Length
        host.activeTabId := host.tabOrder[1]

    GetEmbedRect(host, &areaX, &areaY, &areaW, &areaH)

    ; Position content area border (1px frame, content inset by 1px)
    if host.HasProp("contentBorderTop") && host.contentBorderTop {
        host.contentBorderTop.Move(areaX, areaY, areaW, 1)
        host.contentBorderBottom.Move(areaX, areaY + areaH - 1, areaW, 1)
        host.contentBorderLeft.Move(areaX, areaY, 1, areaH)
        host.contentBorderRight.Move(areaX + areaW - 1, areaY, 1, areaH)
        hasTabs := host.activeTabId != ""
        host.contentBorderTop.Visible := hasTabs
        host.contentBorderBottom.Visible := hasTabs
        host.contentBorderLeft.Visible := hasTabs
        host.contentBorderRight.Visible := hasTabs
    }

    if host.activeTabId = "" {
        DrawTabBar(host)
        return
    }

    ; Inset content area by 1px for border
    areaX += 1
    areaY += 1
    areaW -= 2
    areaH -= 2

    ; Show active tab first, then hide inactive ones.
    ; This order prevents a blank-background flash: the new content is visible before the old one disappears.
    activeHwnd := ""
    for tabId in host.tabOrder {
        if !host.tabRecords.Has(tabId)
            continue
        if tabId != host.activeTabId
            continue
        record := host.tabRecords[tabId]
        if !IsWindowExists(record.contentHwnd)
            continue
        ; Position without SWP_NOCOPYBITS — avoids erasing pixels before the window repaints.
        flags := SWP_FRAMECHANGED | SWP_NOZORDER | SWP_NOACTIVATE
        DllCall("SetWindowPos", "ptr", record.contentHwnd, "ptr", 0
            , "int", areaX, "int", areaY, "int", areaW, "int", areaH
            , "uint", flags)
        DllCall("ShowWindow", "ptr", record.contentHwnd, "int", SW_SHOWNOACTIVATE)
        activeHwnd := record.contentHwnd
        break
    }
    ; Now hide all inactive tabs
    for tabId in host.tabOrder {
        if !host.tabRecords.Has(tabId)
            continue
        if tabId = host.activeTabId
            continue
        record := host.tabRecords[tabId]
        if IsWindowExists(record.contentHwnd)
            DllCall("ShowWindow", "ptr", record.contentHwnd, "int", SW_HIDE)
    }
    ; Immediate redraw so embedded content paints (RDW_NOERASE in RedrawAnyWindow avoids flash).
    if activeHwnd
        RedrawAnyWindow(activeHwnd)

    DrawTabBar(host)
    host.lastRefreshActiveTabId := host.activeTabId
    ; Single deferred redraw at 50ms for slow apps (e.g. WPF) that need time to settle.
    ; Debounced: cancel previous so we don't stack redraws when switching tabs quickly.
    if host.activeTabId != "" {
        if host.HasProp("deferredRepaintFn") && host.deferredRepaintFn
            SetTimer(host.deferredRepaintFn, 0)
        host.deferredRepaintFn := DeferredRepaintCheck.Bind(host)
        SetTimer(host.deferredRepaintFn, -50)
    }
    } finally {
        DllCall("LockWindowUpdate", "Ptr", 0)
    }
}

DrawTabBar(host) {
    global g_HostPadding, g_TabGap, g_MinTabWidth, g_MaxTabWidth, g_TabCornerRadius, g_TabSeparatorWidth
    global g_HeaderHeight, g_TabHeight, g_CloseButtonWidth, g_PopoutButtonWidth
    global g_TabBarAlignment, g_TabBarOffsetY, g_TabIndicatorHeight
    global g_TabPosition, g_UseCustomTitleBar, g_TitleBarHeight
    global g_ThemeTabBarBg, g_ThemeTabActiveBg, g_ThemeTabActiveText
    global g_ThemeTabInactiveBg, g_ThemeTabInactiveBgHover, g_ThemeTabInactiveText
    global g_ThemeTabIndicatorColor, g_ThemeTabSeparatorColor, g_ThemeContentBorder, g_ThemeIconFont, g_ThemeIconFontSize
    global g_ThemeFontNameTab, g_ThemeFontSize, g_ThemeIconColor, g_ShowTabNumbers, g_TabTitleAlignH, g_TabTitleAlignV
    global g_IconClose, g_IconPopout, g_IconMerge

    alignHVal := (g_TabTitleAlignH = "left") ? 0 : (g_TabTitleAlignH = "right") ? 2 : 1
    alignVVal := (g_TabTitleAlignV = "top") ? 0 : (g_TabTitleAlignV = "bottom") ? 2 : 1

    if !host.HasProp("tabCanvas") || !host.tabCanvas
        return
    if !host.hwnd || !IsWindowExists(host.hwnd)
        return

    w := GetClientWidth(host.hwnd)
    h := GetClientHeight(host.hwnd)
    if !w || !h
        return

    tabBarW := w
    tabBarH := g_HeaderHeight
    tabBarY := g_UseCustomTitleBar ? g_TitleBarHeight : 0
    if g_TabPosition = "bottom"
        tabBarY := h - g_HeaderHeight

    host.tabCanvas.Move(0, tabBarY, tabBarW, tabBarH)

    GdipCreateOffscreenBitmap(tabBarW, tabBarH, &pBitmap, &pGraphics)
    if !pBitmap || !pGraphics {
        if pGraphics
            DllCall("gdiplus\GdipDeleteGraphics", "UPtr", pGraphics)
        if pBitmap
            DllCall("gdiplus\GdipDisposeImage", "UPtr", pBitmap)
        SetTimer(() => DrawTabBar(host), -16)
        return
    }

    ; Fill background
    pBgBrush := 0
    DllCall("gdiplus\GdipCreateSolidFill", "UInt", HexToARGB(g_ThemeTabBarBg),
        "UPtr*", &pBgBrush)
    DllCall("gdiplus\GdipFillRectangleI", "UPtr", pGraphics, "UPtr", pBgBrush,
        "Int", 0, "Int", 0, "Int", tabBarW, "Int", tabBarH)
    DllCall("gdiplus\GdipDeleteBrush", "UPtr", pBgBrush)

    tabCount := host.tabOrder.Length
    if tabCount = 0 {
        ApplyBitmapToCanvas(host.tabCanvas, pBitmap, pGraphics)
        return
    }

    arrowW := 24
    usableWidth := Max(200, tabBarW - (g_HostPadding * 2))
    tabWidth := Floor((usableWidth - ((tabCount - 1) * g_TabGap)) / tabCount)
    tabWidth := Max(g_MinTabWidth, Min(g_MaxTabWidth, tabWidth))
    ; Overflow: if tabs exceed available width, switch to scroll mode
    totalW := tabCount * tabWidth + (tabCount - 1) * g_TabGap
    needScroll := totalW > usableWidth
    effectivePopoutW := g_ShowPopoutButton ? g_PopoutButtonWidth : 0
    effectiveCloseW  := g_ShowCloseButton  ? g_CloseButtonWidth  : 0
    titleWidth := tabWidth - effectivePopoutW - effectiveCloseW

    ; Clamp scroll offset and compute how many tabs fit
    if needScroll {
        visibleCount := Max(1, Floor((usableWidth - arrowW * 2) / (tabWidth + g_TabGap)))
        host.tabScrollMax := Max(0, tabCount - visibleCount)
        host.tabScrollOffset := Max(0, Min(host.tabScrollOffset, host.tabScrollMax))
        drawStart := host.tabScrollOffset + 1   ; 1-based index into host.tabOrder
        drawEnd   := Min(tabCount, host.tabScrollOffset + visibleCount)
        x := g_HostPadding + arrowW
    } else {
        host.tabScrollMax := 0
        host.tabScrollOffset := 0
        drawStart := 1
        drawEnd   := tabCount
        x := g_HostPadding
    }

    if g_TabBarOffsetY >= 0
        tabOffsetY := g_TabBarOffsetY
    else {
        align := StrLower(g_TabBarAlignment)
        if align = "top"
            tabOffsetY := 0
        else if align = "bottom"
            tabOffsetY := tabBarH - g_TabHeight
        else
            tabOffsetY := (tabBarH - g_TabHeight) // 2
    }

    for i, tabId in host.tabOrder {
        if needScroll && (i < drawStart || i > drawEnd)
            continue
        isActive  := (tabId = host.activeTabId)
        isHovered := (tabId = host.tabHoveredId)

        bgColor := isActive  ? HexToARGB(g_ThemeTabActiveBg)
                 : isHovered ? HexToARGB(g_ThemeTabInactiveBgHover)
                 :             HexToARGB(g_ThemeTabInactiveBg)
        fgColor := isActive  ? HexToARGB(g_ThemeTabActiveText)
                 :             HexToARGB(g_ThemeTabInactiveText)
        iconColor := isActive ? HexToARGB(g_ThemeTabActiveText)
                   :            HexToARGB(g_ThemeIconColor)

        ; Tab background with rounded corners (radius 5)
        GdipFillRoundRect(pGraphics, x, tabOffsetY, tabWidth, g_TabHeight, g_TabCornerRadius, bgColor)

        ; Active indicator strip
        if isActive && g_TabIndicatorHeight > 0 {
            indicColor := HexToARGB(g_ThemeTabIndicatorColor != ""
                ? g_ThemeTabIndicatorColor : g_ThemeTabActiveBg)
            indicY := (g_TabPosition = "bottom")
                ? tabOffsetY
                : tabOffsetY + g_TabHeight - g_TabIndicatorHeight
            indicW := Max(1, tabWidth - 8)
            ; Radius must not exceed half the height — GDI+ arc math goes negative otherwise.
            ; Use 0 (sharp rect) when indicator is thinner than 3px.
            indicR := (g_TabIndicatorHeight >= 3) ? 1 : 0
            GdipFillRoundRect(pGraphics, x + 4, indicY, indicW, g_TabIndicatorHeight, indicR, indicColor)
        }

        ; Tab title
        rawTitle := host.tabRecords.Has(tabId)
            ? FilterTitle(host.tabRecords[tabId].title) : "Window"
        if g_ShowTabNumbers
            rawTitle := i ". " rawTitle
        GdipDrawStringSimple(pGraphics, rawTitle,
            x, tabOffsetY, titleWidth, g_TabHeight,
            fgColor, g_ThemeFontNameTab, g_ThemeFontSize, isActive,
            true, true, alignHVal, alignVVal)

        ; Popout / merge icon (only drawn if ShowPopoutButton)
        if g_ShowPopoutButton {
            iconText := host.isPopout ? g_IconMerge : g_IconPopout
            GdipDrawStringSimple(pGraphics, iconText,
                x + titleWidth, tabOffsetY, g_PopoutButtonWidth, g_TabHeight,
                iconColor, g_ThemeIconFont, g_ThemeIconFontSize, false)
        }

        ; Close icon (only drawn if ShowCloseButton)
        if g_ShowCloseButton {
            GdipDrawStringSimple(pGraphics, g_IconClose,
                x + titleWidth + g_PopoutButtonWidth, tabOffsetY,
                g_CloseButtonWidth, g_TabHeight,
                iconColor, g_ThemeIconFont, g_ThemeIconFontSize, false)
        }

        x += tabWidth + g_TabGap
    }

    ; Vertical separators between tabs
    if g_TabSeparatorWidth > 0 {
        sepColor := g_ThemeTabSeparatorColor != "" ? HexToARGB(g_ThemeTabSeparatorColor) : HexToARGB(g_ThemeContentBorder)
        pSepBrush := 0
        DllCall("gdiplus\GdipCreateSolidFill", "UInt", sepColor, "UPtr*", &pSepBrush)
        sepX := needScroll ? g_HostPadding + arrowW : g_HostPadding
        for i, tabId in host.tabOrder {
            if needScroll && (i < drawStart || i > drawEnd)
                continue
            if (needScroll ? i < drawEnd : i < tabCount) {
                DllCall("gdiplus\GdipFillRectangleI", "UPtr", pGraphics, "UPtr", pSepBrush,
                    "Int", sepX + tabWidth, "Int", tabOffsetY, "Int", g_TabSeparatorWidth, "Int", g_TabHeight)
            }
            sepX += tabWidth + g_TabGap
        }
        DllCall("gdiplus\GdipDeleteBrush", "UPtr", pSepBrush)
    }

    ; Draw left scroll arrow only when can scroll left (tabScrollOffset > 0)
    if needScroll && host.tabScrollOffset > 0 {
        arrowColor := HexToARGB(g_ThemeTabActiveText)
        GdipDrawStringSimple(pGraphics, Chr(0xE76B),
            g_HostPadding, tabOffsetY, arrowW, g_TabHeight,
            arrowColor, g_ThemeIconFont, g_ThemeIconFontSize, false)
    }

    ; Draw right scroll arrow only when can scroll right (tabScrollOffset < tabScrollMax)
    if needScroll && host.tabScrollOffset < host.tabScrollMax {
        arrowColor := HexToARGB(g_ThemeTabActiveText)
        GdipDrawStringSimple(pGraphics, Chr(0xE76C),
            g_HostPadding + arrowW + (drawEnd - drawStart + 1) * (tabWidth + g_TabGap) - g_TabGap,
            tabOffsetY, arrowW, g_TabHeight,
            arrowColor, g_ThemeIconFont, g_ThemeIconFontSize, false)
    }

    ApplyBitmapToCanvas(host.tabCanvas, pBitmap, pGraphics)
}

; Updates host window title with tab count and active tab name.
UpdateHostTitle(host) {
    global g_HostTitle

    if !host || !host.gui
        return

    liveCount := GetLiveTabCount(host)
    if host.isPopout
        suffix := " (popped out)"
    else
        suffix := ""
    if (host.activeTabId != "") && host.tabRecords.Has(host.activeTabId)
        title := g_HostTitle . " (" liveCount ") - " . host.tabRecords[host.activeTabId].title . suffix
    else
        title := g_HostTitle . " (" liveCount ")" . suffix
    if title = (host.HasProp("lastRefreshTitle") ? host.lastRefreshTitle : "")
        return
    host.lastRefreshTitle := title
    host.gui.Title := title
    if host.HasProp("titleText") && host.titleText
        host.titleText.Text := title
    UpdateHostIcon(host)
}

; Computes content area rect (x, y, w, h) for embedded windows.
GetEmbedRect(host, &x, &y, &w, &h) {
    global g_HostWidth, g_HostHeight, g_HostPadding, g_HostPaddingBottom, g_HeaderHeight
    global g_UseCustomTitleBar, g_TitleBarHeight, g_TabPosition

    padBottom := (g_HostPaddingBottom >= 0) ? g_HostPaddingBottom : g_HostPadding
    x := g_HostPadding
    customTitleH := g_UseCustomTitleBar ? g_TitleBarHeight : 0

    if g_TabPosition = "bottom"
        y := customTitleH + g_HostPadding
    else
        y := customTitleH + g_HeaderHeight + g_HostPadding

    if host.hwnd && IsWindowExists(host.hwnd) {
        try {
            WinGetClientPos(,, &clientW, &clientH, "ahk_id " host.hwnd)
            w := Max(200, clientW - (g_HostPadding * 2))
            if g_TabPosition = "bottom"
                h := Max(140, clientH - y - g_HeaderHeight - padBottom)
            else
                h := Max(140, clientH - y - padBottom)
            return
        }
    }

    w := Max(200, g_HostWidth - (g_HostPadding * 2))
    if g_TabPosition = "bottom"
        h := Max(140, g_HostHeight - y - g_HeaderHeight - padBottom)
    else
        h := Max(140, g_HostHeight - y - padBottom)
}

; Exit handler: restores all tabs, clears state.
CleanupAll(*) {
    global g_MainHost, g_PopoutHosts, g_IsCleaningUp, g_PendingCandidates, g_WatchdogTimerActive
    global g_WinEventHooks, g_WinEventHookCallback

    if g_IsCleaningUp
        return

    g_IsCleaningUp := true
    g_WatchdogTimerActive := false
    SetTimer(WatchdogCheck, 0)
    SetTimer(TooltipCheckTimer, 0)

    if IsSet(g_WinEventHooks) {
        for hook in g_WinEventHooks
            DllCall("UnhookWinEvent", "Ptr", hook)
        g_WinEventHooks := []
    }

    for host in GetAllHosts() {
        for tabId in host.tabOrder.Clone()
            RemoveTrackedTab(host, tabId, true)
    }

    g_PendingCandidates := Map()
    ; Release cached GDI+ font objects
    global g_CachedFontFamily, g_CachedFont, g_CachedStringFormat, g_GdipToken
    for _, pFamily in g_CachedFontFamily
        DllCall("gdiplus\GdipDeleteFontFamily", "UPtr", pFamily)
    g_CachedFontFamily := Map()
    for _, pFont in g_CachedFont
        DllCall("gdiplus\GdipDeleteFont", "UPtr", pFont)
    g_CachedFont := Map()
    if IsObject(g_CachedStringFormat) {
        for _, pFmt in g_CachedStringFormat
            DllCall("gdiplus\GdipDeleteStringFormat", "UPtr", pFmt)
        g_CachedStringFormat := 0
    }
    if g_GdipToken {
        DllCall("gdiplus\GdiplusShutdown", "UPtr", g_GdipToken)
        g_GdipToken := 0
    }
    g_IsCleaningUp := false
}

; Writes discovered candidate windows to discovery.txt (when DebugDiscovery=1).
DumpDiscoveryDebug() {
    global g_DebugLogPath

    discovered := DiscoverCandidateWindows()
    text := "Timestamp: " FormatTime(, "yyyy-MM-dd HH:mm:ss") "`r`n"
    text .= "Discovered windows: " discovered.Length "`r`n`r`n"

    for candidate in discovered {
        text .= "Tab ID: " candidate.id "`r`n"
        text .= candidate.hierarchySummary "`r`n"
        text .= "------------------------------`r`n"
    }

    FileDelete(g_DebugLogPath)
    FileAppend(text, g_DebugLogPath, "UTF-8")
    MsgBox("Wrote discovery info to:`n" g_DebugLogPath, "StackTabs Debug")
}

; Builds debug string describing top window, owner chain, content, and descendants.
DescribeWindowHierarchy(topHwnd, contentHwnd) {
    lines := []
    lines.Push("Top: " . DescribeSingleWindow(topHwnd))
    lines.Push("Root owner: " GetRootOwner(topHwnd))

    ownerChain := GetOwnerChain(topHwnd)
    if ownerChain.Length {
        ownerText := ""
        for idx, hwnd in ownerChain {
            if idx > 1
                ownerText .= " -> "
            ownerText .= hwnd
        }
        lines.Push("Owner chain: " ownerText)
    }

    if contentHwnd != topHwnd
        lines.Push("Chosen content: " . DescribeSingleWindow(contentHwnd))

    descendants := GetDescendantWindows(topHwnd)
    if descendants.Length {
        lines.Push("Descendants:")
        for hwnd in descendants
            lines.Push("  " . DescribeSingleWindow(hwnd))
    }

    return JoinLines(lines)
}

; Returns one-line debug description of a window (class, process, title, size, style).
DescribeSingleWindow(hwnd) {
    if !IsWindowExists(hwnd)
        return hwnd " [missing]"

    title := SafeWinGetTitle(hwnd)
    className := GetWindowClassName(hwnd)
    processName := SafeWinGetProcessName(hwnd)
    parent := DllCall("GetParent", "ptr", hwnd, "ptr")
    owner := GetWindowOwner(hwnd)
    visible := DllCall("IsWindowVisible", "ptr", hwnd) ? "visible" : "hidden"
    style := Format("0x{:08X}", GetWindowLongPtrValue(hwnd, -16))
    exStyle := Format("0x{:08X}", GetWindowLongPtrValue(hwnd, -20))

    try WinGetPos(, , &w, &h, "ahk_id " hwnd)
    catch {
        w := 0
        h := 0
    }

    return hwnd " [" className "] (" processName ") title='" title "' parent=" parent " owner=" owner " size=" w "x" h " " visible " style=" style " ex=" exStyle
}

; Returns array of owner hwnds (top-level owner chain, max 12).
GetOwnerChain(hwnd) {
    chain := []
    seen := Map()
    current := hwnd

    loop 12 {
        owner := GetWindowOwner(current)
        if !owner
            break
        if seen.Has(owner "")
            break

        seen[owner ""] := true
        chain.Push(owner)
        current := owner
    }

    return chain
}

; Returns title from top window, or content window if top has no title.
GetPreferredTabTitle(record) {
    title := SafeWinGetTitle(record.topHwnd)
    if title != ""
        return title
    return SafeWinGetTitle(record.contentHwnd)
}

; Forces window to redraw (invalidates and updates).
RedrawAnyWindow(hwnd) {
    if !hwnd || !IsWindowExists(hwnd)
        return

    flags := 0x0001 | 0x0004 | 0x0080 | 0x0100 | 0x0400  ; INVALIDATE|ERASENOW|UPDATENOW|ALLCHILDREN|FRAME
    DllCall("RedrawWindow", "ptr", hwnd, "ptr", 0, "ptr", 0, "uint", flags)
    DllCall("UpdateWindow", "ptr", hwnd)
}

; Timer callback: redraws active tab content and host after layout change.
DeferredRepaintCheck(host, *) {
    if host.HasProp("deferredRepaintFn")
        host.deferredRepaintFn := ""
    if !host || !host.hwnd || !IsWindowExists(host.hwnd)
        return
    if host.activeTabId != "" && host.tabRecords.Has(host.activeTabId) {
        record := host.tabRecords[host.activeTabId]
        if IsWindowExists(record.contentHwnd)
            RedrawAnyWindow(record.contentHwnd)
    }
    RedrawAnyWindow(host.hwnd)
}

; ============ ICON ============

; Sets host taskbar icon to active tab's icon with theme-colored badge.
UpdateHostIcon(host) {
    global g_ThemeTabActiveBg

    lastId := host.HasProp("lastIconTabId") ? host.lastIconTabId : ""
    if lastId = host.activeTabId
        return
    host.lastIconTabId := host.activeTabId

    if host.HasProp("iconHandle") && host.iconHandle {
        DllCall("DestroyIcon", "ptr", host.iconHandle)
        host.iconHandle := 0
    }

    if !host.activeTabId || !host.tabRecords.Has(host.activeTabId)
        return

    record := host.tabRecords[host.activeTabId]
    hSource := GetWindowBestIcon(record.topHwnd)
    hBadged := CreateBadgedIcon(hSource, g_ThemeTabActiveBg)
    if !hBadged
        return

    host.iconHandle := hBadged
    ; SendMessage fails for hidden windows; enable detection so host can be found when ShowOnlyWhenTabs
    prev := DetectHiddenWindows(true)
    try {
        SendMessage(0x0080, 0, hBadged,, "ahk_id " host.hwnd)  ; WM_SETICON ICON_SMALL
        SendMessage(0x0080, 1, hBadged,, "ahk_id " host.hwnd)  ; WM_SETICON ICON_BIG
    } finally {
        DetectHiddenWindows(prev)
    }
}

; Gets best available icon from window (ICON_SMALL2, ICON_BIG, ICON_SMALL, or class default).
GetWindowBestIcon(hwnd) {
    hIcon := 0
    try hIcon := SendMessage(0x7F, 2, 0,, "ahk_id " hwnd)  ; WM_GETICON ICON_SMALL2
    if !hIcon
        try hIcon := SendMessage(0x7F, 1, 0,, "ahk_id " hwnd)  ; WM_GETICON ICON_BIG
    if !hIcon
        try hIcon := SendMessage(0x7F, 0, 0,, "ahk_id " hwnd)  ; WM_GETICON ICON_SMALL
    if !hIcon {
        fnName := (A_PtrSize = 8) ? "GetClassLongPtr" : "GetClassLong"
        hIcon := DllCall(fnName, "ptr", hwnd, "int", -14, "ptr")  ; GCLP_HICON
        if !hIcon
            hIcon := DllCall(fnName, "ptr", hwnd, "int", -34, "ptr")  ; GCLP_HICONSM
    }
    return hIcon
}

; Draws source icon on 32x32 bitmap with theme-colored badge in corner; returns HICON (caller must DestroyIcon).
CreateBadgedIcon(hSourceIcon, badgeHex) {
    sz := 32

    hScreenDC := DllCall("GetDC", "ptr", 0, "ptr")
    hMemDC    := DllCall("CreateCompatibleDC", "ptr", hScreenDC, "ptr")
    hBmp      := DllCall("CreateCompatibleBitmap", "ptr", hScreenDC, "int", sz, "int", sz, "ptr")
    DllCall("ReleaseDC", "ptr", 0, "ptr", hScreenDC)
    hOldBmp   := DllCall("SelectObject", "ptr", hMemDC, "ptr", hBmp, "ptr")

    ; Black background so icon transparency renders naturally
    RECT := Buffer(16, 0)
    NumPut("int", sz, RECT, 8), NumPut("int", sz, RECT, 12)
    DllCall("FillRect", "ptr", hMemDC, "ptr", RECT, "ptr", DllCall("GetStockObject", "int", 4, "ptr"))

    if hSourceIcon
        DllCall("DrawIconEx", "ptr", hMemDC, "int", 0, "int", 0, "ptr", hSourceIcon
            , "int", sz, "int", sz, "uint", 0, "ptr", 0, "uint", 3)

    ; Parse badge color RRGGBB -> COLORREF 0x00BBGGRR
    r := Integer("0x" SubStr(badgeHex, 1, 2))
    g := Integer("0x" SubStr(badgeHex, 3, 2))
    b := Integer("0x" SubStr(badgeHex, 5, 2))
    badgeColor := r | (g << 8) | (b << 16)

    ; White border circle (1 px larger all around)
    bsz := 10
    hNullPen    := DllCall("GetStockObject", "int", 8, "ptr")  ; NULL_PEN
    hWhiteBrush := DllCall("CreateSolidBrush", "uint", 0xFFFFFF, "ptr")
    hOldPen   := DllCall("SelectObject", "ptr", hMemDC, "ptr", hNullPen, "ptr")
    hOldBrush := DllCall("SelectObject", "ptr", hMemDC, "ptr", hWhiteBrush, "ptr")
    DllCall("Ellipse", "ptr", hMemDC, "int", sz-bsz-1, "int", sz-bsz-1, "int", sz, "int", sz)
    DllCall("DeleteObject", "ptr", hWhiteBrush)

    ; Accent-coloured fill
    hBadgeBrush := DllCall("CreateSolidBrush", "uint", badgeColor, "ptr")
    DllCall("SelectObject", "ptr", hMemDC, "ptr", hBadgeBrush, "ptr")
    DllCall("Ellipse", "ptr", hMemDC, "int", sz-bsz, "int", sz-bsz, "int", sz, "int", sz)
    DllCall("DeleteObject", "ptr", hBadgeBrush)

    DllCall("SelectObject", "ptr", hMemDC, "ptr", hOldBrush, "ptr")
    DllCall("SelectObject", "ptr", hMemDC, "ptr", hOldPen, "ptr")

    ; Monochrome AND-mask â€” all black = fully opaque
    hMaskDC  := DllCall("CreateCompatibleDC", "ptr", 0, "ptr")
    hMaskBmp := DllCall("CreateBitmap", "int", sz, "int", sz, "uint", 1, "uint", 1, "ptr", 0, "ptr")
    hOldMask := DllCall("SelectObject", "ptr", hMaskDC, "ptr", hMaskBmp, "ptr")
    DllCall("PatBlt", "ptr", hMaskDC, "int", 0, "int", 0, "int", sz, "int", sz, "uint", 0x00000042)
    DllCall("SelectObject", "ptr", hMaskDC, "ptr", hOldMask, "ptr")
    DllCall("DeleteDC", "ptr", hMaskDC)

    ; ICONINFO layout: fIcon(4), xHotspot(4), yHotspot(4), [pad4 on x64], hbmMask(ptr), hbmColor(ptr)
    maskOff  := (A_PtrSize = 8) ? 16 : 12
    colorOff := (A_PtrSize = 8) ? 24 : 16
    ICONINFO  := Buffer((A_PtrSize = 8) ? 32 : 20, 0)
    NumPut("int", 1,        ICONINFO, 0)
    NumPut("ptr", hMaskBmp, ICONINFO, maskOff)
    NumPut("ptr", hBmp,     ICONINFO, colorOff)
    hIcon := DllCall("CreateIconIndirect", "ptr", ICONINFO, "ptr")

    DllCall("SelectObject", "ptr", hMemDC, "ptr", hOldBmp, "ptr")
    DllCall("DeleteDC", "ptr", hMemDC)
    ; Always delete the GDI objects — CreateIconIndirect makes internal copies
    DllCall("DeleteObject", "ptr", hMaskBmp)
    if !hIcon
        DllCall("DeleteObject", "ptr", hBmp)
    return hIcon
}

; Appends text to discovery.txt when DebugDiscovery=1.
AppendDebugLog(text, critical := false) {
    global g_DebugLogPath, g_DebugDiscovery
    if !g_DebugDiscovery && !critical
        return
    FileAppend("[" FormatTime(, "yyyy-MM-dd HH:mm:ss") "]`r`n" text, g_DebugLogPath, "UTF-8")
}

; Joins array of strings with CRLF.
JoinLines(lines) {
    text := ""
    for idx, line in lines {
        if idx > 1
            text .= "`r`n"
        text .= line
    }
    return text
}

; Lowercases, trims, and collapses whitespace for title comparison.
NormalizeTitle(title) {
    normalized := Trim(StrLower(title))
    normalized := RegExReplace(normalized, "\s+", " ")
    return normalized
}

; Removes TitleFilters patterns from title for display.
FilterTitle(title) {
    global g_TitleStripPatterns
    for pattern in g_TitleStripPatterns
        title := RegExReplace(title, pattern, "")
    return Trim(title)
}

; Truncates title to maxLen with "..." suffix.
ShortTitle(title, maxLen := 28) {
    if StrLen(title) <= maxLen
        return title
    return SubStr(title, 1, maxLen - 1) . "..."
}

; Selects the tab at the given 1-based index (used by Ctrl+1–9 hotkeys).
SelectTabByIndex(host, idx) {
    if !host || idx < 1 || idx > host.tabOrder.Length
        return
    SelectTab(host, host.tabOrder[idx])
}

; Cycles to next/previous tab (direction 1 or -1).
CycleTabs(host, direction) {
    count := host.tabOrder.Length
    if count < 2
        return

    currentIndex := 0
    for idx, tabId in host.tabOrder {
        if tabId = host.activeTabId {
            currentIndex := idx
            break
        }
    }

    if currentIndex = 0
        currentIndex := 1

    nextIndex := currentIndex + direction
    if nextIndex > count
        nextIndex := 1
    else if nextIndex < 1
        nextIndex := count

    SelectTab(host, host.tabOrder[nextIndex])
}

; Returns true if the active window belongs to a StackTabs host.
StackTabsHostIsActive(*) {
    return !!GetActiveStackTabsHost()
}

; Returns the host that owns the currently active window, or "".
GetActiveStackTabsHost() {
    try
        activeHwnd := WinGetID("A")
    catch
        return ""
    if !activeHwnd
        return ""
    return GetHostForHwnd(activeHwnd)
}

; WinGetTitle wrapper that returns "" on error.
SafeWinGetTitle(hwnd) {
    try return WinGetTitle("ahk_id " hwnd)
    catch
        return ""
}

; WinGetProcessName wrapper that returns "" on error.
SafeWinGetProcessName(hwnd) {
    try return WinGetProcessName("ahk_id " hwnd)
    catch
        return ""
}

; Returns the window class name via GetClassName.
GetWindowClassName(hwnd) {
    buf := Buffer(512, 0)
    DllCall("GetClassName", "ptr", hwnd, "ptr", buf, "int", 256)
    return StrGet(buf)
}

; Returns root ancestor (GA_ROOTOWNER) of window.
GetRootOwner(hwnd) {
    return DllCall("GetAncestor", "ptr", hwnd, "uint", 3, "ptr")
}

; Returns owner window (GW_OWNER).
GetWindowOwner(hwnd) {
    return DllCall("GetWindow", "ptr", hwnd, "uint", 4, "ptr")
}

; Gets window long (GWL_*); works on both 32- and 64-bit.
GetWindowLongPtrValue(hwnd, index) {
    if A_PtrSize = 8
        return DllCall("GetWindowLongPtr", "ptr", hwnd, "int", index, "ptr")
    return DllCall("GetWindowLong", "ptr", hwnd, "int", index, "ptr")
}

; Sets window long (GWL_*); works on both 32- and 64-bit.
SetWindowLongPtrValue(hwnd, index, value) {
    if A_PtrSize = 8
        return DllCall("SetWindowLongPtr", "ptr", hwnd, "int", index, "ptr", value, "ptr")
    return DllCall("SetWindowLong", "ptr", hwnd, "int", index, "ptr", value, "ptr")
}

GetTabWidthForHost(host) {
    global g_HostPadding, g_TabGap, g_MinTabWidth, g_MaxTabWidth
    tabCount := host.tabOrder.Length
    if tabCount = 0
        return 0
    w := GetClientWidth(host.hwnd)
    usableWidth := Max(200, w - (g_HostPadding * 2))
    tabWidth := Floor((usableWidth - ((tabCount - 1) * g_TabGap)) / tabCount)
    tabWidth := Max(g_MinTabWidth, Min(g_MaxTabWidth, tabWidth))
    return tabWidth
}

GetTabIndexAtMouseX(host, mouseX) {
    global g_HostPadding, g_TabGap
    tabWidth := GetTabWidthForHost(host)
    if tabWidth = 0
        return 0
    tabCount := host.tabOrder.Length
    arrowW := 24
    needScroll := host.HasProp("tabScrollMax") && host.tabScrollMax > 0
    startX    := needScroll ? g_HostPadding + arrowW : g_HostPadding
    startIdx  := needScroll ? host.tabScrollOffset + 1 : 1
    Loop tabCount {
        i := A_Index
        if i < startIdx
            continue
        slotX := startX + (i - startIdx) * (tabWidth + g_TabGap)
        if mouseX >= slotX && mouseX < slotX + tabWidth
            return i
    }
    return 0
}

GetTabZoneAtMouseX(host, mouseX, tabIdx) {
    global g_HostPadding, g_CloseButtonWidth, g_PopoutButtonWidth, g_TabGap
    tabWidth := GetTabWidthForHost(host)
    arrowW := 24
    needScroll := host.HasProp("tabScrollMax") && host.tabScrollMax > 0
    startX   := needScroll ? g_HostPadding + arrowW : g_HostPadding
    startIdx := needScroll ? host.tabScrollOffset + 1 : 1
    slotX := startX + (tabIdx - startIdx) * (tabWidth + g_TabGap)
    effectivePopoutW := g_ShowPopoutButton ? g_PopoutButtonWidth : 0
    effectiveCloseW  := g_ShowCloseButton  ? g_CloseButtonWidth  : 0
    titleWidth := tabWidth - effectiveCloseW - effectivePopoutW
    if mouseX < slotX + titleWidth
        return "title"
    if mouseX < slotX + titleWidth + g_PopoutButtonWidth
        return "popout"
    return "close"
}

; Returns true if the cursor is over any host's tab bar (screen coords).
IsMouseOverAnyTabBar() {
    global g_HeaderHeight, g_UseCustomTitleBar, g_TitleBarHeight, g_TabPosition
    MouseGetPos(&screenX, &screenY)
    pt := Buffer(8, 0)
    for host in GetAllHosts() {
        NumPut("Int", screenX, pt, 0)
        NumPut("Int", screenY, pt, 4)
        DllCall("ScreenToClient", "Ptr", host.hwnd, "Ptr", pt)
        clientX := NumGet(pt, 0, "Int")
        clientY := NumGet(pt, 4, "Int")
        clientW := GetClientWidth(host.hwnd)
        clientH := GetClientHeight(host.hwnd)
        if clientX < 0 || clientY < 0 || clientX >= clientW || clientY >= clientH
            continue
        tabBarY := g_UseCustomTitleBar ? g_TitleBarHeight : 0
        if g_TabPosition = "bottom"
            tabBarY := clientH - g_HeaderHeight
        if clientY >= tabBarY && clientY < tabBarY + g_HeaderHeight
            return true
    }
    return false
}

OnTabCanvasMouseMove(wParam, lParam, msg, hwnd) {
    global g_HeaderHeight, g_UseCustomTitleBar, g_TitleBarHeight, g_TabPosition
    mouseX := lParam & 0xFFFF
    mouseY := (lParam >> 16) & 0xFFFF
    for host in GetAllHosts() {
        if host.hwnd != hwnd && (!host.HasProp("tabCanvas") || host.tabCanvas.Hwnd != hwnd)
            continue
        tabBarY := g_UseCustomTitleBar ? g_TitleBarHeight : 0
        if g_TabPosition = "bottom" {
            clientH := GetClientHeight(host.hwnd)
            tabBarY := clientH - g_HeaderHeight
        }
        checkY := (hwnd = host.tabCanvas.Hwnd) ? 0 : tabBarY
        newHoveredId := ""
        if mouseY >= checkY && mouseY < checkY + g_HeaderHeight {
            tabIdx := GetTabIndexAtMouseX(host, mouseX)
            newHoveredId := (tabIdx > 0) ? host.tabOrder[tabIdx] : ""
        }
        if newHoveredId != host.tabHoveredId {
            host.tabHoveredId := newHoveredId
            DrawTabBar(host)
        }
        if newHoveredId != "" && host.tabRecords.Has(newHoveredId) {
            fullTitle := FilterTitle(host.tabRecords[newHoveredId].title)
            if StrLen(fullTitle) > 28 {
                ToolTip(fullTitle)
                SetTimer(TooltipCheckTimer, 100)
            } else {
                ToolTip()
                SetTimer(TooltipCheckTimer, 0)
            }
        } else {
            ToolTip()
            SetTimer(TooltipCheckTimer, 0)
        }
        return
    }
    ToolTip()
    SetTimer(TooltipCheckTimer, 0)
}

TooltipCheckTimer(*) {
    if !IsMouseOverAnyTabBar() {
        ToolTip()
        SetTimer(TooltipCheckTimer, 0)
    }
}

OnTabCanvasClick(wParam, lParam, msg, hwnd) {
    global g_HeaderHeight, g_UseCustomTitleBar, g_TitleBarHeight, g_TabPosition
    global g_HostPadding, g_TabGap
    arrowW := 24
    mouseX := lParam & 0xFFFF
    mouseY := (lParam >> 16) & 0xFFFF
    for host in GetAllHosts() {
        if host.hwnd != hwnd && (!host.HasProp("tabCanvas") || host.tabCanvas.Hwnd != hwnd)
            continue
        tabBarY := g_UseCustomTitleBar ? g_TitleBarHeight : 0
        if g_TabPosition = "bottom" {
            clientH := GetClientHeight(host.hwnd)
            tabBarY := clientH - g_HeaderHeight
        }
        ; When message is from tab canvas, coords are canvas-relative (tab bar origin = 0,0)
        checkY := (hwnd = host.tabCanvas.Hwnd) ? 0 : tabBarY
        if mouseY < checkY || mouseY >= checkY + g_HeaderHeight
            continue
        ; Scroll arrows — only handle when arrow is visible (can scroll in that direction)
        if host.tabScrollMax > 0 {
            if host.tabScrollOffset > 0 && mouseX >= g_HostPadding && mouseX < g_HostPadding + arrowW {
                host.tabScrollOffset := Max(0, host.tabScrollOffset - 1)
                DrawTabBar(host)
                return
            }
            tabWidth := GetTabWidthForHost(host)
            visibleCount := Max(1, Floor((GetClientWidth(host.hwnd) - (g_HostPadding * 2) - arrowW * 2) / (tabWidth + g_TabGap)))
            rightArrowX := g_HostPadding + arrowW + visibleCount * (tabWidth + g_TabGap) - g_TabGap
            if host.tabScrollOffset < host.tabScrollMax && mouseX >= rightArrowX && mouseX < rightArrowX + arrowW {
                host.tabScrollOffset := Min(host.tabScrollMax, host.tabScrollOffset + 1)
                DrawTabBar(host)
                return
            }
        }
        tabIdx := GetTabIndexAtMouseX(host, mouseX)
        if tabIdx = 0 || tabIdx > host.tabOrder.Length
            return
        tabId := host.tabOrder[tabIdx]
        zone := GetTabZoneAtMouseX(host, mouseX, tabIdx)
        if zone = "close"
            CloseTab(host, tabId)
        else if zone = "popout"
            if host.isPopout
                MergeBackTab(host, tabId)
            else
                PopOutTab(host, tabId)
        else
            SelectTab(host, tabId)
        return
    }
}

OnTabCanvasMidClick(wParam, lParam, msg, hwnd) {
    global g_HeaderHeight, g_UseCustomTitleBar, g_TitleBarHeight, g_TabPosition
    mouseX := lParam & 0xFFFF
    mouseY := (lParam >> 16) & 0xFFFF
    for host in GetAllHosts() {
        if host.hwnd != hwnd && (!host.HasProp("tabCanvas") || host.tabCanvas.Hwnd != hwnd)
            continue
        tabBarY := g_UseCustomTitleBar ? g_TitleBarHeight : 0
        if g_TabPosition = "bottom" {
            clientH := GetClientHeight(host.hwnd)
            tabBarY := clientH - g_HeaderHeight
        }
        checkY := (hwnd = host.tabCanvas.Hwnd) ? 0 : tabBarY
        if mouseY < checkY || mouseY >= checkY + g_HeaderHeight
            return
        tabIdx := GetTabIndexAtMouseX(host, mouseX)
        if tabIdx > 0 && tabIdx <= host.tabOrder.Length
            CloseTab(host, host.tabOrder[tabIdx])
        return
    }
}

OnTabCanvasRightClick(wParam, lParam, msg, hwnd) {
    global g_HeaderHeight, g_UseCustomTitleBar, g_TitleBarHeight, g_TabPosition
    mouseX := lParam & 0xFFFF
    mouseY := (lParam >> 16) & 0xFFFF
    for host in GetAllHosts() {
        if host.hwnd != hwnd && (!host.HasProp("tabCanvas") || host.tabCanvas.Hwnd != hwnd)
            continue
        tabBarY := g_UseCustomTitleBar ? g_TitleBarHeight : 0
        if g_TabPosition = "bottom" {
            clientH := GetClientHeight(host.hwnd)
            tabBarY := clientH - g_HeaderHeight
        }
        checkY := (hwnd = host.tabCanvas.Hwnd) ? 0 : tabBarY
        if mouseY < checkY || mouseY >= checkY + g_HeaderHeight
            return
        tabIdx := GetTabIndexAtMouseX(host, mouseX)
        if tabIdx = 0 || tabIdx > host.tabOrder.Length
            return
        tabId := host.tabOrder[tabIdx]
        if !host.tabRecords.Has(tabId)
            return
        record := host.tabRecords[tabId]
        m := Menu()
        h := host
        t := tabId
        ttl := record.title
        if host.isPopout
            m.Add("Merge to Main", ((a, b, *) => MergeBackTab(a, b)).Bind(h, t))
        else
            m.Add("Pop Out", ((a, b, *) => PopOutTab(a, b)).Bind(h, t))
        m.Add("Copy Title", ((s, *) => (A_Clipboard := s)).Bind(ttl))
        m.Add()
        m.Add("Close Tab", ((a, b, *) => CloseTab(a, b)).Bind(h, t))
        m.Show()
        return
    }
}

OnMessage(0x0200, OnTabCanvasMouseMove)
OnMessage(0x0201, OnTabCanvasClick)
OnMessage(0x0204, OnTabCanvasRightClick)
OnMessage(0x0207, OnTabCanvasMidClick)

; === GDI+ helpers (gdiplus.dll) ===

GdiplusStartup() {
    input := Buffer(16, 0)
    NumPut("UInt", 1, input, 0)
    token := 0
    DllCall("gdiplus\GdiplusStartup", "UPtr*", &token, "Ptr", input, "Ptr", 0)
    return token
}

HexToARGB(hex) {
    return Integer("0xFF" . hex)
}

GdipCreateOffscreenBitmap(w, h, &pBitmap, &pGraphics) {
    pBitmap := 0
    pGraphics := 0
    DllCall("gdiplus\GdipCreateBitmapFromScan0",
        "Int", w, "Int", h, "Int", 0, "Int", 0x26200A, "Ptr", 0, "UPtr*", &pBitmap)
    DllCall("gdiplus\GdipGetImageGraphicsContext",
        "UPtr", pBitmap, "UPtr*", &pGraphics)
    DllCall("gdiplus\GdipSetSmoothingMode", "UPtr", pGraphics, "Int", 4)
    DllCall("gdiplus\GdipSetTextRenderingHint", "UPtr", pGraphics, "Int", 5)
}

GdipCleanupBitmap(pBitmap, pGraphics) {
    DllCall("gdiplus\GdipDeleteGraphics", "UPtr", pGraphics)
    DllCall("gdiplus\GdipDisposeImage", "UPtr", pBitmap)
}

GdipBitmapToHBITMAP(pBitmap) {
    hBmp := 0
    DllCall("gdiplus\GdipCreateHBITMAPFromBitmap",
        "UPtr", pBitmap, "UPtr*", &hBmp, "UInt", 0)
    return hBmp
}

GdipFillRoundRect(pGraphics, x, y, w, h, radius, argbColor) {
    if radius <= 0 {
        pBrush := 0
        DllCall("gdiplus\GdipCreateSolidFill", "UInt", argbColor, "UPtr*", &pBrush)
        DllCall("gdiplus\GdipFillRectangleI", "UPtr", pGraphics, "UPtr", pBrush,
            "Int", x, "Int", y, "Int", w, "Int", h)
        DllCall("gdiplus\GdipDeleteBrush", "UPtr", pBrush)
        return
    }
    pPath := 0
    DllCall("gdiplus\GdipCreatePath", "Int", 0, "UPtr*", &pPath)
    DllCall("gdiplus\GdipAddPathArcI", "UPtr", pPath,
        "Int", x, "Int", y, "Int", radius*2, "Int", radius*2,
        "Float", 180.0, "Float", 90.0)
    DllCall("gdiplus\GdipAddPathArcI", "UPtr", pPath,
        "Int", x+w-(radius*2), "Int", y, "Int", radius*2, "Int", radius*2,
        "Float", 270.0, "Float", 90.0)
    DllCall("gdiplus\GdipAddPathArcI", "UPtr", pPath,
        "Int", x+w-(radius*2), "Int", y+h-(radius*2), "Int", radius*2, "Int", radius*2,
        "Float", 0.0, "Float", 90.0)
    DllCall("gdiplus\GdipAddPathArcI", "UPtr", pPath,
        "Int", x, "Int", y+h-(radius*2), "Int", radius*2, "Int", radius*2,
        "Float", 90.0, "Float", 90.0)
    DllCall("gdiplus\GdipClosePathFigure", "UPtr", pPath)
    pBrush := 0
    DllCall("gdiplus\GdipCreateSolidFill", "UInt", argbColor, "UPtr*", &pBrush)
    DllCall("gdiplus\GdipFillPath", "UPtr", pGraphics,
        "UPtr", pBrush, "UPtr", pPath)
    DllCall("gdiplus\GdipDeleteBrush", "UPtr", pBrush)
    DllCall("gdiplus\GdipDeletePath", "UPtr", pPath)
}

GdipDrawStringSimple(pGraphics, text, x, y, w, h, argbColor, fontFamilyName, fontSize, bold, noWrap := true, ellipsis := true, alignH := 1, alignV := 1) {
    global g_CachedFontFamily, g_CachedFont, g_CachedStringFormat

    ; Build cache keys
    familyKey := fontFamilyName
    fontKey := fontFamilyName "|" fontSize "|" (bold ? 1 : 0)

    ; Create or reuse font family
    if !g_CachedFontFamily.Has(familyKey) {
        pFamily := 0
        DllCall("gdiplus\GdipCreateFontFamilyFromName",
            "Str", fontFamilyName, "Ptr", 0, "UPtr*", &pFamily)
        if !pFamily
            return
        g_CachedFontFamily[familyKey] := pFamily
    }
    pFamily := g_CachedFontFamily[familyKey]

    ; Create or reuse font
    if !g_CachedFont.Has(fontKey) {
        pFont := 0
        DllCall("gdiplus\GdipCreateFont",
            "UPtr", pFamily, "Float", Float(fontSize),
            "Int", bold ? 1 : 0, "Int", 3, "UPtr*", &pFont)
        if !pFont
            return
        g_CachedFont[fontKey] := pFont
    }
    pFont := g_CachedFont[fontKey]

    ; GDI+ StringAlignment: 0=Near/left, 1=Center, 2=Far/right
    fmtKey := (noWrap ? "1" : "0") "|" (ellipsis ? "1" : "0") "|" alignH "|" alignV
    if !g_CachedStringFormat || !g_CachedStringFormat.Has(fmtKey) {
        if !IsObject(g_CachedStringFormat)
            g_CachedStringFormat := Map()
        pFormat := 0
        DllCall("gdiplus\GdipCreateStringFormat", "Int", 0, "Int", 0, "UPtr*", &pFormat)
        if !pFormat
            return
        DllCall("gdiplus\GdipSetStringFormatAlign", "UPtr", pFormat, "Int", alignH)
        DllCall("gdiplus\GdipSetStringFormatLineAlign", "UPtr", pFormat, "Int", alignV)
        if noWrap
            DllCall("gdiplus\GdipSetStringFormatFlags", "UPtr", pFormat, "Int", 0x00001000)
        if ellipsis
            DllCall("gdiplus\GdipSetStringFormatTrimming", "UPtr", pFormat, "Int", 5)
        g_CachedStringFormat[fmtKey] := pFormat
    }

    pBrush := 0
    DllCall("gdiplus\GdipCreateSolidFill", "UInt", argbColor, "UPtr*", &pBrush)
    if !pBrush
        return

    rect := Buffer(16, 0)
    NumPut("Float", Float(x), rect, 0)
    NumPut("Float", Float(y), rect, 4)
    NumPut("Float", Float(w), rect, 8)
    NumPut("Float", Float(h), rect, 12)

    DllCall("gdiplus\GdipDrawString",
        "UPtr", pGraphics, "Str", text, "Int", -1,
        "UPtr", pFont, "Ptr", rect, "UPtr", g_CachedStringFormat[fmtKey], "UPtr", pBrush)

    DllCall("gdiplus\GdipDeleteBrush", "UPtr", pBrush)
}

ApplyBitmapToCanvas(canvasCtrl, pBitmap, pGraphics) {
    hBmp := GdipBitmapToHBITMAP(pBitmap)
    GdipCleanupBitmap(pBitmap, pGraphics)
    oldBmp := SendMessage(0x172, 0, hBmp,, "ahk_id " canvasCtrl.Hwnd)
    if oldBmp
        DllCall("DeleteObject", "UPtr", oldBmp)
}

OnTabCanvasMouseWheel(wParam, lParam, msg, hwnd) {
    global g_HeaderHeight, g_UseCustomTitleBar, g_TitleBarHeight, g_TabPosition
    ; WM_MOUSEWHEEL goes to the focus window (usually embedded content), not the window under cursor.
    ; Use cursor position (screen coords in lParam) and check if it's over any host's tab bar.
    screenX := (lParam & 0xFFFF) | ((lParam & 0x8000) ? 0xFFFF0000 : 0)
    screenY := ((lParam >> 16) & 0xFFFF) | (((lParam >> 16) & 0x8000) ? 0xFFFF0000 : 0)
    delta := (wParam >> 16) > 32767 ? -1 : 1   ; WHEEL_DELTA sign
    pt := Buffer(8, 0)
    for host in GetAllHosts() {
        NumPut("Int", screenX, pt, 0)
        NumPut("Int", screenY, pt, 4)
        DllCall("ScreenToClient", "Ptr", host.hwnd, "Ptr", pt)
        clientX := NumGet(pt, 0, "Int")
        clientY := NumGet(pt, 4, "Int")
        tabBarY := g_UseCustomTitleBar ? g_TitleBarHeight : 0
        if g_TabPosition = "bottom" {
            clientH := GetClientHeight(host.hwnd)
            tabBarY := clientH - g_HeaderHeight
        }
        if clientY < tabBarY || clientY >= tabBarY + g_HeaderHeight
            continue
        if host.tabScrollMax > 0 {
            host.tabScrollOffset := Max(0, Min(host.tabScrollMax, host.tabScrollOffset - delta))
            DrawTabBar(host)
        }
        return
    }
}
OnMessage(0x020A, OnTabCanvasMouseWheel)

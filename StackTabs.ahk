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
Config := {
    ConfigPath:     A_ScriptDir "\config.ini",
    ConfigExample:  A_ScriptDir "\config.ini.example",
    ThemesDir:      A_ScriptDir "\themes",
    DebugLogPath:   A_ScriptDir "\discovery.txt",
    DebugDiscovery: false,  ; when true, AppendDebugLog writes to discovery.txt on new candidates

    ; Window title patterns: Match1/Match2/... in config.ini. Window must contain at least one.
    WindowTitleMatches: [],

    ; Optional EXE filter. Leave blank to match any process.
    TargetExe: "",

    ; Shell Hook + Slow Sweep: event-driven discovery; fallback scan interval.
    SlowSweepInterval: 10000,
    StackDelayMs: 30,        ; Minimum wait before stacking; title-stability check provides additional protection
    StackSwitchDelayMs: 150, ; Delay before switching to newly stacked tab; lets content load to reduce glitch
    WatchdogMaxMs: 1500,
    TabDisappearGraceMs: 300,

    ; Host window defaults.
    HostTitle: "StackTabs",
    HostWidth: 1200,
    HostHeight: 800,
    HostMinWidth: 700,
    HostMinHeight: 500,
    HostPadding: 8,
    HostPaddingBottom: -1,   ; -1 = use HostPadding; >=0 = use this for bottom padding
    HeaderHeight: 36,
    TabGap: 6,
    MinTabWidth: 120,
    MaxTabWidth: 240,
    TabHeight: 30,
    TabSlotMax: 50,
    CloseButtonWidth: 22,
    PopoutButtonWidth: 22,
    TabBarAlignment: "center",  ; top, center, or bottom — tabs aligned within the tab bar
    TabBarOffsetY: -1,         ; legacy: -1 = use alignment; >=0 = use as pixel offset (overrides alignment)
    TabPosition: "top",    ; "top" or "bottom"
    TabIndicatorHeight: 3,  ; height in px of the active-tab indicator strip; 0 to disable
    TabCornerRadius: 5,
    ActiveTabStyle: "full",  ; "full" = active tab has different bg; "indicator" = only indicator strip, same bg as inactive
    ShowOnlyWhenTabs: true,  ; when true, show host only when 1+ tabs; hide to tray when 0 (default). Set to 0 to always show host.

    ; === THEME (loaded from themes\ folder; dark.ini is the default and fallback) ===
    ThemeTabIndicatorColor: "",   ; set by LoadThemeFromFile; defaults to TabActiveBg
    ThemeIconFont: "",   ; auto-detected at startup; override in theme file with IconFont=
    ActiveThemeFile: "dark.ini",   ; overridden by ThemeFile= in config.ini
    UseCustomTitleBar: false,
    TitleBarHeight: 28,

    ; Icon codepoints from Segoe Fluent Icons / Segoe MDL2 Assets (same PUA values)
    IconClose: Chr(0xe894),
    IconPopout: Chr(0xE8A7),
    IconMerge: Chr(0xe944),

    ; === TITLE FILTERS ===
    ; Strip patterns loaded from [TitleFilters] Strip1/Strip2/... in StackTabs.ini.
    TitleStripPatterns: [],
    ; Maximum characters shown in a tab label. 9999 = no cap (fully dynamic from tab width); set lower to force shorter titles.
    TabTitleMaxLen: 9999,
    ; Max lines in tab label (1 = single line with ellipsis; 2+ = word wrap to multiple lines).
    TabMaxLines: 1,
    ; Tab title alignment: H = left, center, right; V = top, center, bottom (center V = vertically centered in tab).
    TabTitleAlignH: "center",
    TabTitleAlignV: "center",
    ; When true, tab titles are prefixed with their 1-based position: "1. Title", "2. Title", etc.
    ShowTabNumbers: false,
    ShowCloseButton: true,
    ShowPopoutButton: true,
    TabSeparatorWidth: 0,
    ThemeTabSeparatorColor: "",

    ; Theme colors (loaded by LoadThemeFromFile; dark.ini defaults)
    ThemeBackground: "1C1C2E",
    ThemeTabBarBg: "13132A",
    ThemeTabActiveBg: "7B6CF6",
    ThemeTabActiveText: "FFFFFF",
    ThemeTabInactiveBg: "252540",
    ThemeTabInactiveBgHover: "30304E",
    ThemeTabInactiveText: "C5CDF0",
    ThemeIconColor: "6878B0",
    ThemeContentBorder: "35355A",
    ThemeWindowText: "E0E8FF",
    ThemeFontName: "Segoe UI",
    ThemeFontNameTab: "Segoe UI Semibold",
    ThemeFontSize: 9,
    ThemeIconFontSize: 16
}

; ============ STATE ============
State := {
    MainHost: "",             ; HostInstance for main window
    PopoutHosts: [],          ; array of HostInstance for popped-out windows
    AllHostsCache: [],
    HostByHwnd: Map(),        ; hwnd string -> host (O(1) lookup)
    PendingCandidates: Map(),   ; tabId -> {firstSeen, candidate} (main host only)
    WatchdogTimerActive: false,
    WatchdogInterval: 50,
    IsCleaningUp: false,
    ShellHookMsg: 0,           ; set at startup
    WinEventHooks: [],
    WinEventHookCallback: 0,   ; set at startup
    GdipToken: 0,              ; set at startup
    CachedFontFamily: Map(),
    CachedFont: Map(),
    CachedStringFormat: 0,
    GdipShutdownPending: false,
    LastHookEventTick: 0,
    SwitcherGui: "",
    SwitcherAllTabs: [],
    SwitcherVisible: [],
    SwitcherSelVisIdx: 0,
    SwitcherCards: []
}

; Loads config.ini into globals; migrates from StackTabs.ini if needed.
LoadConfigFromIni() {
    ; Migrate from StackTabs.ini if config.ini doesn't exist
    if !FileExist(Config.ConfigPath) && FileExist(A_ScriptDir "\StackTabs.ini")
        FileCopy(A_ScriptDir "\StackTabs.ini", Config.ConfigPath)
    if !FileExist(Config.ConfigPath)
        return
    iniPath := Config.ConfigPath
    ; Read theme first so it's applied even if the try block below throws (e.g. invalid Layout values)
    Config.ActiveThemeFile := Trim(IniRead(iniPath, "Theme", "ThemeFile", "dark.ini"))
    try {
        Config.TargetExe := IniRead(iniPath, "General", "TargetExe", Config.TargetExe)
        Config.SlowSweepInterval := Integer(IniRead(iniPath, "General", "SlowSweepInterval", Config.SlowSweepInterval))
        Config.StackDelayMs := Integer(IniRead(iniPath, "General", "StackDelayMs", Config.StackDelayMs))
        Config.StackSwitchDelayMs := Integer(IniRead(iniPath, "General", "StackSwitchDelayMs", Config.StackSwitchDelayMs))
        Config.WatchdogMaxMs := Integer(IniRead(iniPath, "General", "WatchdogMaxMs", Config.WatchdogMaxMs))
        Config.TabDisappearGraceMs := Integer(IniRead(iniPath, "General", "TabDisappearGraceMs", Config.TabDisappearGraceMs))
        Config.DebugDiscovery := (IniRead(iniPath, "General", "DebugDiscovery", "0") = "1")
        Config.HostTitle := IniRead(iniPath, "Layout", "HostTitle", Config.HostTitle)
        Config.HostWidth := Integer(IniRead(iniPath, "Layout", "HostWidth", Config.HostWidth))
        Config.HostHeight := Integer(IniRead(iniPath, "Layout", "HostHeight", Config.HostHeight))
        Config.HostMinWidth := Integer(IniRead(iniPath, "Layout", "HostMinWidth", Config.HostMinWidth))
        Config.HostMinHeight := Integer(IniRead(iniPath, "Layout", "HostMinHeight", Config.HostMinHeight))
        Config.HostPadding := Integer(IniRead(iniPath, "Layout", "HostPadding", Config.HostPadding))
        Config.HostPaddingBottom := Integer(IniRead(iniPath, "Layout", "HostPaddingBottom", "-1"))
        Config.HeaderHeight := Integer(IniRead(iniPath, "Layout", "HeaderHeight", Config.HeaderHeight))
        Config.TabGap := Integer(IniRead(iniPath, "Layout", "TabGap", Config.TabGap))
        Config.MinTabWidth := Integer(IniRead(iniPath, "Layout", "MinTabWidth", Config.MinTabWidth))
        Config.MaxTabWidth := Integer(IniRead(iniPath, "Layout", "MaxTabWidth", Config.MaxTabWidth))
        Config.TabHeight := Integer(IniRead(iniPath, "Layout", "TabHeight", Config.TabHeight))
        Config.TabSlotMax := Integer(IniRead(iniPath, "Layout", "TabSlotMax", Config.TabSlotMax))
        Config.CloseButtonWidth := Integer(IniRead(iniPath, "Layout", "CloseButtonWidth", Config.CloseButtonWidth))
        Config.PopoutButtonWidth := Integer(IniRead(iniPath, "Layout", "PopoutButtonWidth", Config.PopoutButtonWidth))
        rawAlign := IniRead(iniPath, "Layout", "TabBarAlignment", "")
        Config.TabBarAlignment := (rawAlign != "") ? Trim(rawAlign) : "center"
        Config.TabBarOffsetY := Integer(IniRead(iniPath, "Layout", "TabBarOffsetY", "-1"))  ; legacy: -1 = use alignment
        Config.TabTitleMaxLen := Integer(Trim(StrSplit(IniRead(iniPath, "Layout", "TabTitleMaxLen", "9999"), ";")[1]))
        Config.TabMaxLines := Max(1, Integer(Trim(StrSplit(IniRead(iniPath, "Layout", "TabMaxLines", Config.TabMaxLines), ";")[1])))
        Config.TabTitleAlignH := Trim(StrSplit(IniRead(iniPath, "Layout", "TabTitleAlignH", Config.TabTitleAlignH), ";")[1])
        Config.TabTitleAlignV := Trim(StrSplit(IniRead(iniPath, "Layout", "TabTitleAlignV", Config.TabTitleAlignV), ";")[1])
        Config.ShowTabNumbers := (Trim(StrSplit(IniRead(iniPath, "Layout", "ShowTabNumbers", "0"), ";")[1]) = "1")
        Config.ShowCloseButton := (Trim(StrSplit(IniRead(iniPath, "Layout", "ShowCloseButton", "1"), ";")[1]) = "1")
        Config.ShowPopoutButton := (Trim(StrSplit(IniRead(iniPath, "Layout", "ShowPopoutButton", "1"), ";")[1]) = "1")
        Config.TabPosition := IniRead(iniPath, "Layout", "TabPosition", "top")
        Config.TabIndicatorHeight := Integer(IniRead(iniPath, "Layout", "TabIndicatorHeight", "3"))
        Config.TabCornerRadius := Integer(IniRead(iniPath, "Layout", "TabCornerRadius", "5"))
        Config.TabSeparatorWidth := Integer(IniRead(iniPath, "Layout", "TabSeparatorWidth", "0"))
        Config.ActiveTabStyle := Trim(IniRead(iniPath, "Layout", "ActiveTabStyle", "full"))
        ; ShowOnlyWhenTabs: show host only when 1+ tabs; hide to tray when 0 (default). Fallback for old config keys.
        rawVal := IniRead(iniPath, "Layout", "ShowOnlyWhenTabs", IniRead(iniPath, "Layout", "KeepHostAlive", IniRead(iniPath, "Layout", "HideHostWhenEmpty", "1")))
        ; Strip inline comment (; ...) and trim so "1   ; comment" parses as 1
        Config.ShowOnlyWhenTabs := (Trim(StrSplit(rawVal, ";")[1]) = "1")
        Config.UseCustomTitleBar := (IniRead(iniPath, "Layout", "UseCustomTitleBar", "0") = "1")
        Config.TitleBarHeight := Integer(IniRead(iniPath, "Layout", "TitleBarHeight", "28"))
    }
    ; Load match patterns from [General] Match1/Match2/...
    Config.WindowTitleMatches := []
    i := 1
    loop {
        val := IniRead(iniPath, "General", "Match" i, "")
        if val = ""
            break
        Config.WindowTitleMatches.Push(val)
        i++
    }
    ; No fallback: require Match1/Match2/... or WindowTitleMatch to be configured
    if Config.WindowTitleMatches.Length = 0 {
        fallback := Trim(IniRead(iniPath, "General", "WindowTitleMatch", ""))
        if fallback != ""
            Config.WindowTitleMatches.Push(fallback)
    }
    ; Load strip patterns from [TitleFilters] section (Strip1, Strip2, ...)
    Config.TitleStripPatterns := []
    i := 1
    loop {
        val := IniRead(iniPath, "TitleFilters", "Strip" i, "")
        if val = ""
            break
        Config.TitleStripPatterns.Push(val)
        i++
    }
}

; Loads theme colors and layout overrides from an .ini file; falls back to dark.ini if missing.
LoadThemeFromFile(themePath) {
    ; Fall back to dark.ini if the requested theme file doesn't exist
    if !FileExist(themePath)
        themePath := Config.ThemesDir "\dark.ini"
    ; Use dark.ini values as fallbacks for missing keys (partial theme files)
    Config.ThemeBackground         := IniRead(themePath, "Theme", "Background",           "1C1C2E")
    Config.ThemeTabBarBg           := IniRead(themePath, "Theme", "TabBarBg",             "13132A")
    Config.ThemeTabActiveBg        := IniRead(themePath, "Theme", "TabActiveBg",          "7B6CF6")
    Config.ThemeTabActiveText      := IniRead(themePath, "Theme", "TabActiveText",        "FFFFFF")
    ; TabIndicatorColor: color of the active-tab indicator strip. Defaults to TabActiveBg.
    Config.ThemeTabIndicatorColor   := IniRead(themePath, "Theme", "TabIndicatorColor",    Config.ThemeTabActiveBg)
    Config.ThemeTabInactiveBg      := IniRead(themePath, "Theme", "TabInactiveBg",        "252540")
    Config.ThemeTabInactiveBgHover := IniRead(themePath, "Theme", "TabInactiveBgHover",   "30304E")
    Config.ThemeTabInactiveText    := IniRead(themePath, "Theme", "TabInactiveText",      "C5CDF0")
    Config.ThemeIconColor          := IniRead(themePath, "Theme", "IconColor",            "6878B0")
    Config.ThemeContentBorder      := IniRead(themePath, "Theme", "ContentBorder",        "35355A")
    Config.ThemeWindowText         := IniRead(themePath, "Theme", "WindowText",           "E0E8FF")
    Config.ThemeFontName           := IniRead(themePath, "Theme", "FontName",             "Segoe UI")
    Config.ThemeFontNameTab        := IniRead(themePath, "Theme", "FontNameTab",          "Segoe UI Semibold")
    Config.ThemeFontSize           := Integer(IniRead(themePath, "Theme", "FontSize",     "9"))
    Config.ThemeIconFont           := IniRead(themePath, "Theme", "IconFont",             "")
    Config.ThemeIconFontSize       := Integer(IniRead(themePath, "Theme", "IconFontSize",   "16"))
    ; Optional layout overrides â€” only applied if the theme file includes a [Layout] section
    Config.HostPadding        := Integer(IniRead(themePath, "Layout", "HostPadding",        String(Config.HostPadding)))
    Config.HostPaddingBottom  := Integer(IniRead(themePath, "Layout", "HostPaddingBottom",  String(Config.HostPaddingBottom)))
    Config.HeaderHeight       := Integer(IniRead(themePath, "Layout", "HeaderHeight",       String(Config.HeaderHeight)))
    Config.TabGap             := Integer(IniRead(themePath, "Layout", "TabGap",             String(Config.TabGap)))
    Config.MinTabWidth        := Integer(IniRead(themePath, "Layout", "MinTabWidth",        String(Config.MinTabWidth)))
    Config.MaxTabWidth        := Integer(IniRead(themePath, "Layout", "MaxTabWidth",        String(Config.MaxTabWidth)))
    Config.TabHeight          := Integer(IniRead(themePath, "Layout", "TabHeight",          String(Config.TabHeight)))
    Config.CloseButtonWidth   := Integer(IniRead(themePath, "Layout", "CloseButtonWidth",   String(Config.CloseButtonWidth)))
    Config.PopoutButtonWidth  := Integer(IniRead(themePath, "Layout", "PopoutButtonWidth",  String(Config.PopoutButtonWidth)))
    ; TabBarAlignment: when theme doesn't specify, read from config so we don't carry over previous theme's value
    rawAlign := IniRead(themePath, "Layout", "TabBarAlignment", "")
    Config.TabBarAlignment := (rawAlign != "") ? Trim(rawAlign) : IniRead(Config.ConfigPath, "Layout", "TabBarAlignment", "center")
    rawOffset := IniRead(themePath, "Layout", "TabBarOffsetY", "")
    if rawOffset != ""
        Config.TabBarOffsetY := Integer(rawOffset)
    Config.TabIndicatorHeight := Integer(IniRead(themePath, "Layout", "TabIndicatorHeight", String(Config.TabIndicatorHeight)))
    Config.TabCornerRadius := Integer(IniRead(themePath, "Layout", "TabCornerRadius", String(Config.TabCornerRadius)))
    Config.TabSeparatorWidth := Integer(IniRead(themePath, "Layout", "TabSeparatorWidth", String(Config.TabSeparatorWidth)))
    Config.ThemeTabSeparatorColor := IniRead(themePath, "Theme", "TabSeparatorColor", "")
    ; TabPosition: when theme doesn't specify, read from config (not Config.TabPosition) so we don't carry over
    ; the previous theme's value when switching themes
    rawPos := IniRead(themePath, "Layout", "TabPosition", "")
    Config.TabPosition := (rawPos != "") ? Trim(rawPos) : IniRead(Config.ConfigPath, "Layout", "TabPosition", "top")
    rawStyle := IniRead(themePath, "Layout", "ActiveTabStyle", "")
    Config.ActiveTabStyle := (rawStyle != "") ? Trim(rawStyle) : "full"
    rawTabMaxLines := IniRead(themePath, "Layout", "TabMaxLines", "")
    if rawTabMaxLines != "" {
        parts := StrSplit(rawTabMaxLines, ";")
        if parts.Length
            Config.TabMaxLines := Max(1, Integer(Trim(parts[1])))
    }
    rawTabTitleMaxLen := IniRead(themePath, "Layout", "TabTitleMaxLen", "")
    if rawTabTitleMaxLen != "" {
        parts := StrSplit(rawTabTitleMaxLen, ";")
        if parts.Length
            Config.TabTitleMaxLen := Integer(Trim(parts[1]))
    }
    rawAlignH := IniRead(themePath, "Layout", "TabTitleAlignH", "")
    if rawAlignH != "" {
        parts := StrSplit(rawAlignH, ";")
        if parts.Length
            Config.TabTitleAlignH := Trim(parts[1])
    }
    rawAlignV := IniRead(themePath, "Layout", "TabTitleAlignV", "")
    if rawAlignV != "" {
        parts := StrSplit(rawAlignV, ";")
        if parts.Length
            Config.TabTitleAlignV := Trim(parts[1])
    }
    rawShowNums := IniRead(themePath, "Layout", "ShowTabNumbers", "")
    if rawShowNums != "" {
        parts := StrSplit(rawShowNums, ";")
        if parts.Length
            Config.ShowTabNumbers := (Trim(parts[1]) = "1")
    } else {
        cfgVal := IniRead(Config.ConfigPath, "Layout", "ShowTabNumbers", "0")
        parts := StrSplit(cfgVal, ";")
        Config.ShowTabNumbers := (parts.Length && Trim(parts[1]) = "1")
    }
    rawShowClose := IniRead(themePath, "Layout", "ShowCloseButton", "")
    if rawShowClose != "" {
        parts := StrSplit(rawShowClose, ";")
        if parts.Length
            Config.ShowCloseButton := (Trim(parts[1]) = "1")
    } else {
        cfgVal := IniRead(Config.ConfigPath, "Layout", "ShowCloseButton", "1")
        parts := StrSplit(cfgVal, ";")
        Config.ShowCloseButton := (parts.Length && Trim(parts[1]) = "1")
    }
    rawShowPopout := IniRead(themePath, "Layout", "ShowPopoutButton", "")
    if rawShowPopout != "" {
        parts := StrSplit(rawShowPopout, ";")
        if parts.Length
            Config.ShowPopoutButton := (Trim(parts[1]) = "1")
    } else {
        cfgVal := IniRead(Config.ConfigPath, "Layout", "ShowPopoutButton", "1")
        parts := StrSplit(cfgVal, ";")
        Config.ShowPopoutButton := (parts.Length && Trim(parts[1]) = "1")
    }
}


; Auto-detects Segoe Fluent Icons or falls back to Segoe MDL2 Assets for icon glyphs.
DetectIconFont() {
    if Config.ThemeIconFont != ""
        return
    try {
        Loop Reg, "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts", "V" {
            if InStr(A_LoopRegName, "Segoe Fluent Icons") {
                Config.ThemeIconFont := "Segoe Fluent Icons"
                return
            }
        }
    }
    Config.ThemeIconFont := "Segoe MDL2 Assets"
}

; Builds tray menu with theme submenu, themes folder link, and Exit.
BuildTrayMenu() {
    A_TrayMenu.Delete()
    themeSubMenu := Menu()
    themesDir := Config.ThemesDir
    if DirExist(themesDir) {
        ; Built-in themes (themes\*.ini)
        Loop Files, themesDir "\*.ini" {
            fileName := A_LoopFileName
            displayName := ThemeDisplayName(fileName)
            themeSubMenu.Add(displayName, ThemeMenuHandler.Bind(fileName))
            if (Trim(fileName) = Trim(Config.ActiveThemeFile))
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
                if (Trim(fileName) = Trim(Config.ActiveThemeFile))
                    try themeSubMenu.Check(displayName)
            }
        }
    }
    A_TrayMenu.Add("Theme", themeSubMenu)
    A_TrayMenu.Add()
    A_TrayMenu.Add("Open themes folder", (*) => Run(Config.ThemesDir))
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
    if !FileExist(Config.ConfigPath) {
        if FileExist(A_ScriptDir "\StackTabs.ini")
            FileCopy(A_ScriptDir "\StackTabs.ini", Config.ConfigPath)
        else if FileExist(Config.ConfigExample)
            FileCopy(Config.ConfigExample, Config.ConfigPath)
    }
    themePath := Config.ThemesDir "\" themeFileName
    if !FileExist(themePath) {
        MsgBox("Theme file not found: " themePath, "StackTabs", "Icon!")
        return
    }
    IniWrite(themeFileName, Config.ConfigPath, "Theme", "ThemeFile")
    ; Free cached GDI+ font objects from the previous theme before loading the new one.
    ; Without this, every theme switch leaks font family and font handles permanently.
    for _, pFamily in State.CachedFontFamily
        DllCall("gdiplus\GdipDeleteFontFamily", "UPtr", pFamily)
    State.CachedFontFamily := Map()
    for _, pFont in State.CachedFont
        DllCall("gdiplus\GdipDeleteFont", "UPtr", pFont)
    State.CachedFont := Map()
    if IsObject(State.CachedStringFormat) {
        for _, pFmt in State.CachedStringFormat
            DllCall("gdiplus\GdipDeleteStringFormat", "UPtr", pFmt)
        State.CachedStringFormat := Map()
    }
    Config.ActiveThemeFile := themeFileName
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
    if !host || !host.gui || !IsWindowExists(host.hwnd)
        return
    host.gui.BackColor := Config.ThemeBackground
    host.gui.SetFont("s" Config.ThemeFontSize " c" Config.ThemeWindowText, Config.ThemeFontName)
    if host.HasProp("tabBarBg") && host.tabBarBg
        host.tabBarBg.Opt("Background0x" Config.ThemeTabBarBg)
    if host.HasProp("contentBorderTop") && host.contentBorderTop {
        host.contentBorderTop.Opt("Background0x" Config.ThemeContentBorder)
        host.contentBorderBottom.Opt("Background0x" Config.ThemeContentBorder)
        host.contentBorderLeft.Opt("Background0x" Config.ThemeContentBorder)
        host.contentBorderRight.Opt("Background0x" Config.ThemeContentBorder)
    }
    if host.HasProp("titleBarBg") && host.titleBarBg {
        host.titleBarBg.Opt("Background0x" Config.ThemeTabBarBg)
        host.titleCloseBtn.Opt("Background0x" Config.ThemeTabBarBg " c" Config.ThemeIconColor)
        if Config.UseCustomTitleBar && host.hwnd {
            try {
                rgb := Integer("0x" Config.ThemeTabBarBg)
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

LoadConfigFromIni()
if Config.WindowTitleMatches.Length = 0 {
    cfgPath := FileExist(Config.ConfigPath) ? Config.ConfigPath : Config.ConfigExample
    MsgBox("No window match patterns configured.`n`n"
        . "Add Match1=, Match2=, etc. in config.ini under [General].`n"
        . "Each value is a substring to match in window titles (e.g. Match1=PowerShell, Match2=Notepad).`n`n"
        . "See config.ini.example for details.", "StackTabs - Configuration Required", "Icon!")
    try Run(cfgPath)
    ExitApp()
}
LoadThemeFromFile(A_ScriptDir "\themes\" Config.ActiveThemeFile)
DetectIconFont()
BuildTrayMenu()
if Config.UseCustomTitleBar
    OnMessage(0x83, OnWmNcCalcSize)
DllCall("LoadLibrary", "Str", "gdiplus", "Ptr")
State.GdipToken := GdiplusStartup()
State.CachedFontFamily := Map()
State.CachedFont := Map()
State.CachedStringFormat := 0
State.GdipShutdownPending := false
BuildHostInstance(false)  ; create main host
; Shell Hook for event-driven window discovery
DllCall("RegisterShellHookWindow", "Ptr", State.MainHost.hwnd)
State.ShellHookMsg := DllCall("RegisterWindowMessage", "Str", "SHELLHOOK", "UInt")
OnMessage(State.ShellHookMsg, OnShellHook)

; WinEvent hooks: faster window detection than the Shell Hook for apps (e.g. WPF) that
; register with the taskbar late. One range registration covers the three events we care about:
;   EVENT_OBJECT_SHOW       (0x8002): window becomes visible — fires before taskbar registration
;   EVENT_OBJECT_NAMECHANGE (0x800C): title changes — catches WPF apps that set title after show
;   EVENT_OBJECT_UNCLOAKED  (0x8018): DWM uncloaks window — catches UWP/WinUI apps
; The callback filters to only those three; everything else in the range is discarded cheaply.
; WINEVENT_OUTOFCONTEXT (0): events are queued to this thread's message loop — no injection.
State.WinEventHookCallback := CallbackCreate(WinEventProc, , 7)
State.WinEventHooks := []
; EVENT_OBJECT_SHOW (0x8002): window becomes visible
State.WinEventHooks.Push(DllCall("SetWinEventHook",
    "UInt", 0x8002, "UInt", 0x8002,
    "Ptr",  0,      "Ptr",  State.WinEventHookCallback,
    "UInt", 0,      "UInt", 0, "UInt", 0, "Ptr"))
; EVENT_OBJECT_NAMECHANGE (0x800C): title changed (WPF sets title after show)
State.WinEventHooks.Push(DllCall("SetWinEventHook",
    "UInt", 0x800C, "UInt", 0x800C,
    "Ptr",  0,      "Ptr",  State.WinEventHookCallback,
    "UInt", 0,      "UInt", 0, "UInt", 0, "Ptr"))
; EVENT_OBJECT_UNCLOAKED (0x8018): UWP/WinUI window uncloaks
State.WinEventHooks.Push(DllCall("SetWinEventHook",
    "UInt", 0x8018, "UInt", 0x8018,
    "Ptr",  0,      "Ptr",  State.WinEventHookCallback,
    "UInt", 0,      "UInt", 0, "UInt", 0, "Ptr"))

OnExit(CleanupAll)

RefreshWindows()
; Slow sweep fallback: only does full WinGetList scan when hooks have been
; quiet for 5+ seconds. Primary discovery is via shell hook and WinEvent hooks.
SetTimer(RefreshWindows, Config.SlowSweepInterval)

; ============ TAB SWITCHER OVERLAY ============
; Alt+Shift+F: floating overlay with tab cards and fuzzy search (when StackTabs is focused).
; Type to filter. Arrow keys navigate. Enter switches. Escape closes.

State.SwitcherGui       := ""
State.SwitcherAllTabs   := []
State.SwitcherVisible   := []
State.SwitcherSelVisIdx := 0
State.SwitcherCards     := []

; Tab switcher: Alt+Shift+F when StackTabs host is active
#HotIf StackTabsHostIsActive()
!+f:: {
    if State.SwitcherGui {
        SwitcherClose()
        return
    }
    ; Defer by 80ms so modifiers are physically released before the GUI opens.
    ; Without this the Edit control receives held modifiers = command mode, not typing mode.
    SetTimer(ShowTabSwitcher, -80)
}
#HotIf

ShowTabSwitcher() {
    if State.SwitcherGui {
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
                    title: record.filteredTitle,
                    isActive: (tabId = host.activeTabId)})
        }
    }
    if allTabs.Length = 0
        return

    State.SwitcherAllTabs   := allTabs
    State.SwitcherVisible   := []
    State.SwitcherSelVisIdx := 1
    Loop allTabs.Length
        State.SwitcherVisible.Push(A_Index)
    for idx, item in allTabs {
        if item.isActive {
            State.SwitcherSelVisIdx := idx
            break
        }
    }

    ; Layout — vertical list (command-palette style)
    overlayWidth   := 420
    rowHeight      := 36
    searchBarHeight := 38
    pad            := 16
    listGap        := 8
    listAreaHeight := Min(allTabs.Length * rowHeight, 600)
    overlayHeight  := pad + searchBarHeight + listGap + listAreaHeight + pad

    ; Position: center on StackTabs host window, fallback to screen center
    switcherHost := GetActiveStackTabsHost()
    if switcherHost && IsWindowExists(switcherHost.hwnd) {
        try {
            WinGetPos(&hx, &hy, &hw, &hh, "ahk_id " switcherHost.hwnd)
            ox := hx + (hw - overlayWidth) // 2
            oy := hy + (hh - overlayHeight) // 2
        } catch {
            ox := (A_ScreenWidth - overlayWidth) // 2
            oy := (A_ScreenHeight - overlayHeight) // 2
        }
    } else {
        ox := (A_ScreenWidth - overlayWidth) // 2
        oy := (A_ScreenHeight - overlayHeight) // 2
    }

    State.SwitcherGui           := Gui("+AlwaysOnTop -Caption +ToolWindow", "TabSwitcher")
    State.SwitcherGui.BackColor  := Config.ThemeTabBarBg
    State.SwitcherGui.MarginX    := 0
    State.SwitcherGui.MarginY    := 0
    State.SwitcherCards := []

    ; Search box — cue-banner text via EM_SETCUEBANNER after show
    State.SwitcherGui.SetFont("s" (Config.ThemeFontSize + 1) " c" Config.ThemeWindowText, Config.ThemeFontName)
    searchBox := State.SwitcherGui.Add("Edit",
        "x" pad " y" pad " w" (overlayWidth - pad * 2) " h" searchBarHeight
        " Background" Config.ThemeTabBarBg " c" Config.ThemeWindowText, "")
    searchBox.SetFont("s" (Config.ThemeFontSize + 1) " c" Config.ThemeWindowText, Config.ThemeFontName)

    listAreaY := pad + searchBarHeight + listGap

    Loop allTabs.Length {
        i    := A_Index
        ry   := listAreaY + (i - 1) * rowHeight
        item := allTabs[i]
        isSel := (i = State.SwitcherSelVisIdx)
        bg := isSel ? Config.ThemeTabActiveBg : Config.ThemeTabInactiveBg
        fg := isSel ? Config.ThemeTabActiveText : Config.ThemeTabInactiveText

        row := State.SwitcherGui.Add("Text",
            "x" pad " y" ry " w" (overlayWidth - pad * 2) " h" rowHeight
            " +0x200 Left Background" bg " c" fg,
            "  " ShortTitle(item.title, 55))
        row.SetFont("s" Config.ThemeFontSize " c" fg, Config.ThemeFontNameTab)
        row.tabSwitcherIdx := i
        row.OnEvent("Click", SwitcherCardClick)
        State.SwitcherCards.Push(row)
    }

    State.SwitcherGui.OnEvent("Close", (*) => SwitcherClose())
    State.SwitcherGui.Show("x" ox " y" oy " w" overlayWidth " h" overlayHeight)
    DllCall("dwmapi.dll\DwmSetWindowAttribute",
        "ptr", State.SwitcherGui.Hwnd, "uint", 33, "uint*", 2, "uint", 4)

    ; Cue banner "Search tabs..." inside the edit box
    SendMessage(0x1501, 1, StrPtr("Search tabs..."),, "ahk_id " searchBox.Hwnd)

    searchBox.OnEvent("Change", SwitcherOnSearch)
    OnMessage(0x0100, SwitcherOnKeyDown, 1)
    WinActivate("ahk_id " State.SwitcherGui.Hwnd)
    searchBox.Focus()
}

SwitcherClose() {
    OnMessage(0x0100, SwitcherOnKeyDown, 0)
    if State.SwitcherGui {
        State.SwitcherGui.Destroy()
        State.SwitcherGui := ""
    }
    State.SwitcherAllTabs := []
    State.SwitcherVisible := []
    State.SwitcherCards   := []
}

SwitcherOnSearch(ctrl, *) {
    query := ctrl.Value
    State.SwitcherVisible := []
    for idx, item in State.SwitcherAllTabs {
        if query = "" || InStr(item.title, query, false)
            State.SwitcherVisible.Push(idx)
    }
    State.SwitcherSelVisIdx := State.SwitcherVisible.Length ? 1 : 0
    SwitcherRefreshCards()
}

SwitcherRefreshCards() {
    if !State.SwitcherGui
        return

    visSet := Map()
    for _, tabIdx in State.SwitcherVisible
        visSet[tabIdx] := true
    selTabIdx := (State.SwitcherSelVisIdx >= 1 && State.SwitcherSelVisIdx <= State.SwitcherVisible.Length)
        ? State.SwitcherVisible[State.SwitcherSelVisIdx] : 0

    for idx, row in State.SwitcherCards {
        isVis := visSet.Has(idx)
        isSel := (idx = selTabIdx)
        row.Visible := isVis
        if isVis {
            bg := isSel ? Config.ThemeTabActiveBg : Config.ThemeTabInactiveBg
            fg := isSel ? Config.ThemeTabActiveText : Config.ThemeTabInactiveText
            row.Opt("Background" bg " c" fg)
            row.SetFont("s" Config.ThemeFontSize " c" fg, Config.ThemeFontNameTab)
        }
    }
}

SwitcherCardClick(ctrl, *) {
    idx := ctrl.tabSwitcherIdx
    for _, tabIdx in State.SwitcherVisible {
        if tabIdx = idx {
            SwitcherActivate(tabIdx)
            return
        }
    }
}

SwitcherActivate(tabIdx) {
    if tabIdx < 1 || tabIdx > State.SwitcherAllTabs.Length
        return
    item := State.SwitcherAllTabs[tabIdx]
    SwitcherClose()
    SelectTab(item.host, item.tabId)
    if item.host.hwnd && IsWindowExists(item.host.hwnd)
        WinActivate("ahk_id " item.host.hwnd)
}

SwitcherOnKeyDown(wParam, lParam, msg, hwnd) {
    if !State.SwitcherGui
        return
    if hwnd != State.SwitcherGui.Hwnd {
        parent := DllCall("GetParent", "ptr", hwnd, "ptr")
        if parent != State.SwitcherGui.Hwnd
            return
    }
    count := State.SwitcherVisible.Length
    if wParam = 0x1B {
        SwitcherClose()
        return 0
    }
    if count = 0
        return
    if wParam = 0x0D {
        if State.SwitcherSelVisIdx >= 1 && State.SwitcherSelVisIdx <= count
            SwitcherActivate(State.SwitcherVisible[State.SwitcherSelVisIdx])
        return 0
    }
    if wParam = 0x25 {
        State.SwitcherSelVisIdx := Max(1, State.SwitcherSelVisIdx - 1)
        SwitcherRefreshCards()
        return 0
    }
    if wParam = 0x27 {
        State.SwitcherSelVisIdx := Min(count, State.SwitcherSelVisIdx + 1)
        SwitcherRefreshCards()
        return 0
    }
    if wParam = 0x26 {
        State.SwitcherSelVisIdx := Max(1, State.SwitcherSelVisIdx - 1)
        SwitcherRefreshCards()
        return 0
    }
    if wParam = 0x28 {
        State.SwitcherSelVisIdx := Min(count, State.SwitcherSelVisIdx + 1)
        SwitcherRefreshCards()
        return 0
    }
}

; ============ HOTKEYS ============
; Win+Shift+T: toggle host visibility (hide when active, show when hidden).
#+t:: {

    if !State.MainHost
        return

    if IsWindowExists(State.MainHost.hwnd) && WinActive("ahk_id " State.MainHost.hwnd)
        State.MainHost.gui.Hide()
    else {
        ; When ShowOnlyWhenTabs: only show if we have tabs
        if Config.ShowOnlyWhenTabs && GetLiveTabCount(State.MainHost) = 0
            return
        State.MainHost.gui.Show()
    }
}

; Win+Shift+D: dump discovery debug to disk (only when DebugDiscovery=1).
#+d:: {
    if !Config.DebugDiscovery
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
    if host.isPopout {
        InvalidateHostsCache()
        for tabId in host.tabOrder.Clone()
            RemoveTrackedTab(host, tabId, true)
        for i, h in State.PopoutHosts {
            if h = host {
                State.PopoutHosts.RemoveAt(i)
                break
            }
        }
        if host.hwnd
            State.HostByHwnd.Delete(host.hwnd "")
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
    State.AllHostsCache := []
}

; Returns array of main host plus all popout hosts.
GetAllHosts() {
    if State.AllHostsCache.Length
        return State.AllHostsCache
    hosts := []
    if State.MainHost
        hosts.Push(State.MainHost)
    for h in State.PopoutHosts
        hosts.Push(h)
    State.AllHostsCache := hosts
    return hosts
}

; Finds the host that owns the given hwnd (or its parent chain).
GetHostForHwnd(hwnd) {
    ; 1. Direct O(1) lookup
    if (host := State.HostByHwnd.Get(hwnd "", ""))
        return host
    ; 2. Walk tabRecords for content/top hwnd (embedded windows)
    for host in GetAllHosts() {
        for tabId, record in host.tabRecords {
            if (record.contentHwnd = hwnd || record.topHwnd = hwnd)
                return host
        }
    }
    ; 3. Parent chain as last resort
    current := DllCall("GetParent", "ptr", hwnd, "ptr")
    while current {
        if (host := State.HostByHwnd.Get(current "", ""))
            return host
        for h in GetAllHosts() {
            for tabId, record in h.tabRecords {
                if (record.contentHwnd = current || record.topHwnd = current)
                    return h
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
    host := State.HostByHwnd.Get(hwnd "", "")
    if !host
        return
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

; Creates a new host window (main or popout) with tab bar, content area, and theme.
BuildHostInstance(isPopout := false) {

    host := Object()
    host.isPopout := isPopout
    host.tabRecords := Map()
    host.tabOrder := []
    host.activeTabId := ""
    host.tabHoveredId := ""
    host.tabScrollOffset := 0   ; index of first visible tab (0-based)
    host.tabScrollMax := 0      ; updated by DrawTabBar
    host.isResizing := false

    title := isPopout ? (Config.HostTitle " (popped out)") : Config.HostTitle
    guiOpts := "+Resize +MinSize" Config.HostMinWidth "x" Config.HostMinHeight
    if Config.UseCustomTitleBar
        guiOpts .= " -Caption +Border"
    host.gui := Gui(guiOpts, title)
    host.gui.BackColor := Config.ThemeBackground
    host.gui.MarginX := 0
    host.gui.MarginY := 0
    host.gui.SetFont("s" Config.ThemeFontSize " c" Config.ThemeWindowText, Config.ThemeFontName)
    host.gui.OnEvent("Close", HostGuiClosed.Bind(host))
    host.gui.OnEvent("Size", HostGuiResized.Bind(host))

    tabBarY := 0
    tabBarH := Config.HeaderHeight
    if Config.UseCustomTitleBar {
        tabBarY := Config.TitleBarHeight
        host.titleBarBg := host.gui.Add("Text", "x0 y0 w" Config.HostWidth " h" Config.TitleBarHeight " +0x200 +0x100 Background" Config.ThemeTabBarBg, "")
        host.titleBarBg.OnEvent("Click", TitleBarDragClick.Bind(host))
        host.titleText := host.gui.Add("Text", "x8 y0 w" (Config.HostWidth - 60) " h" Config.TitleBarHeight " +0x200 +0x100 BackgroundTrans", title)
        host.titleText.OnEvent("Click", TitleBarDragClick.Bind(host))
        host.titleCloseBtn := host.gui.Add("Text", "x" (Config.HostWidth - 46) " y0 w46 h" Config.TitleBarHeight " +0x200 +0x100 Center Background" Config.ThemeTabBarBg, Config.IconClose)
        host.titleCloseBtn.SetFont("s" Config.ThemeIconFontSize, Config.ThemeIconFont)
        host.titleCloseBtn.Opt("c" Config.ThemeIconColor)
        host.titleCloseBtn.OnEvent("Click", TitleBarCloseClick.Bind(host))
    }
    host.tabBarBg := host.gui.Add("Text", "x0 y" tabBarY " w" Config.HostWidth " h" tabBarH " Background" Config.ThemeTabBarBg, "")
    host.tabCanvas := host.gui.Add("Pic",
        "x0 y" tabBarY " w" Config.HostWidth " h" tabBarH " +0xE", "")
    host.contentBorderTop := host.gui.Add("Text", "Hidden x0 y0 w0 h1 Background" Config.ThemeContentBorder, "")
    host.contentBorderBottom := host.gui.Add("Text", "Hidden x0 y0 w0 h1 Background" Config.ThemeContentBorder, "")
    host.contentBorderLeft := host.gui.Add("Text", "Hidden x0 y0 w1 h0 Background" Config.ThemeContentBorder, "")
    host.contentBorderRight := host.gui.Add("Text", "Hidden x0 y0 w1 h0 Background" Config.ThemeContentBorder, "")
    host.hwnd := host.gui.Hwnd
    host.clientHwnd := host.hwnd
    State.HostByHwnd[host.hwnd ""] := host
    showOpts := "w" Config.HostWidth " h" Config.HostHeight
    ; Always pass Hide when the window should not be visible yet.
    ; Show() followed immediately by Hide() briefly makes the window visible to the OS,
    ; which causes tiling window managers to tile it as a floating window on first use.
    ; Passing Hide to Show() sets the size without ever showing the window.
    if isPopout || (!isPopout && Config.ShowOnlyWhenTabs)
        showOpts .= " Hide"
    host.gui.Show(showOpts)

    ; Request Windows 11 rounded corners (no-op if already applied by system)
    cornerPref := 2  ; DWM_WCP_ROUND
    DllCall("dwmapi.dll\DwmSetWindowAttribute", "ptr", host.hwnd, "uint", 33, "uint*", cornerPref, "uint", 4)

    ; Match DWM window border to theme (fixes white/accent bar on Windows 11)
    if Config.UseCustomTitleBar && host.hwnd {
        try {
            rgb := Integer("0x" Config.ThemeTabBarBg)
            bgr := ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF)
            DllCall("dwmapi\DwmSetWindowAttribute", "ptr", host.hwnd, "int", 34, "uint*", &bgr, "uint", 4)
        }
    }

    if isPopout
        State.PopoutHosts.Push(host)
    else
        State.MainHost := host
    InvalidateHostsCache()
    return host
}

; Handles host close: restores tabs to main (popout) or hides main host (keeps script running).
HostGuiClosed(host, *) {
    if host.isPopout {
        ; Restore all tabs in this popout to their original parent (release, don't merge)
        for tabId in host.tabOrder.Clone()
            RemoveTrackedTab(host, tabId, true)
        for i, h in State.PopoutHosts {
            if h = host {
                State.PopoutHosts.RemoveAt(i)
                break
            }
        }
        if host.hwnd
            State.HostByHwnd.Delete(host.hwnd "")
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
    if !State.PendingCandidates.Has(candidate.id) {
        ; First time seeing this candidate â€” record the timestamp
        now := A_TickCount
        State.PendingCandidates[candidate.id] := {firstSeen: now, candidate: candidate}
        AppendDebugLog("New candidate (pending watchdog)`r`n" candidate.hierarchySummary "`r`n")
    } else {
        ; Already pending â€” refresh metadata only, preserve firstSeen so the delay is not reset
        State.PendingCandidates[candidate.id].candidate := candidate
    }
    if !State.WatchdogTimerActive {
        State.WatchdogTimerActive := true
        State.WatchdogInterval := 50
        SetTimer(WatchdogCheck, -State.WatchdogInterval)
    }
}

; Processes a single pending candidate: builds fresh candidate, handles duplicate detection/dialog, creates tracked tab.
; Returns the new tab id on success, or "" if skipped/failed.
ProcessPendingCandidate(tabId, pending) {
    ; Re-build candidate from scratch so we use fresh, stable metadata (not the snapshot from creation time)
    freshCandidate := BuildCandidateFromTopWindow(pending.candidate.topHwnd)
    if !IsObject(freshCandidate)
        return ""
    ; Skip if already embedded (could have been stacked by slow sweep between checks)
    if State.MainHost.tabRecords.Has(freshCandidate.id)
        return ""
    ; Duplicate detection: same process + same title = offer to close the older one
    existingTabId := FindTabWithSameTitle(State.MainHost, freshCandidate)
    if existingTabId != "" {
        result := DuplicateConfirmDialog(ShortTitle(freshCandidate.title, 50))
        if (result = "Yes")
            CloseTab(State.MainHost, existingTabId)
    }
    if CreateTrackedTab(State.MainHost, freshCandidate)
        return freshCandidate.id
    return ""
}

; Timer callback: stacks candidates that passed delay + title-stability; removes stale pending.
WatchdogCheck(*) {
    if !State.MainHost || State.PendingCandidates.Count = 0 {
        State.WatchdogTimerActive := false
        SetTimer(WatchdogCheck, 0)
        State.WatchdogInterval := 50   ; reset for next activation
        return
    }
    now := A_TickCount
    toStack := []
    toRemove := []
    for tabId, pending in State.PendingCandidates {
        elapsed := now - pending.firstSeen
        if elapsed >= Config.StackDelayMs && IsReadyToStack(pending.candidate.topHwnd) {
            ; Title-stability check: stack only once the title has been unchanged for two consecutive ticks (~50ms)
            currentTitle := SafeWinGetTitle(pending.candidate.topHwnd)
            if pending.HasProp("lastSeenTitle") && (currentTitle = pending.lastSeenTitle)
                toStack.Push({tabId: tabId, candidate: pending.candidate, firstSeen: pending.firstSeen})
            else {
                pending.lastSeenTitle := currentTitle
            }
        } else if elapsed >= Config.WatchdogMaxMs {
            toRemove.Push(tabId)
        }
    }
    anyStacked := false
    lastStackedTabId := ""
    hadTabsBefore := State.MainHost.tabOrder.Length
    for item in toStack {
        State.PendingCandidates.Delete(item.tabId)
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
            if State.MainHost.HasProp("pendingSwitchTimer") && State.MainHost.pendingSwitchTimer
                SetTimer(State.MainHost.pendingSwitchTimer, 0)
            State.MainHost.pendingSwitchTimer := SwitchToNewTabDelayed.Bind(State.MainHost, lastStackedTabId)
            SetTimer(State.MainHost.pendingSwitchTimer, -Config.StackSwitchDelayMs)
        } else {
            ; First tab: let the window settle hidden, then show. One-shot.
            State.MainHost.activeTabId := lastStackedTabId
            SetTimer(WatchdogPostStackUpdate, -Config.StackSwitchDelayMs)
        }
    }
    for tabId in toRemove
        State.PendingCandidates.Delete(tabId)
    if State.PendingCandidates.Count = 0 {
        State.WatchdogTimerActive := false
        SetTimer(WatchdogCheck, 0)
        State.WatchdogInterval := 50
    } else {
        ; Adaptive back-off: start at 50ms, double each tick up to 400ms.
        ; Resets to 50ms when candidates are cleared.
        State.WatchdogInterval := Min(400, State.WatchdogInterval * 2)
        SetTimer(WatchdogCheck, -State.WatchdogInterval)
    }
}

; Deferred update after stacking: show host if needed, layout tabs, refresh content.
WatchdogPostStackUpdate(*) {
    if !State.MainHost
        return
    ; Only call Show when the window is actually hidden â€” avoids repositioning an already-visible window
    if Config.ShowOnlyWhenTabs && State.MainHost.tabOrder.Length >= 1 {
        if !WinExist("ahk_id " State.MainHost.hwnd)
            State.MainHost.gui.Show()  ; Show() activates by default; intentional for first appearance
        ; If already visible don't force-activate — user may have focus elsewhere
    }
    LayoutTabButtons(State.MainHost)
    ShowOnlyActiveTab(State.MainHost)
    UpdateHostTitle(State.MainHost)
    RedrawAnyWindow(State.MainHost.hwnd)
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
                        UpdateRecordTitleCache(record)
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
    if !State.MainHost || State.IsCleaningUp
        return
    if State.SwitcherGui && hwnd = State.SwitcherGui.Hwnd
        return
    ; Fast HWND check before the expensive candidate build (which walks all descendants).
    ; Covers NAMECHANGE events on already-tracked tabs whose title changed.
    for host in GetAllHosts() {
        for tabId, record in host.tabRecords {
            if (record.topHwnd = hwnd || record.contentHwnd = hwnd)
                return
        }
    }
    for tabId, pending in State.PendingCandidates {
        if pending.candidate.topHwnd = hwnd
            return
    }
    State.LastHookEventTick := A_TickCount
    candidate := BuildCandidateFromTopWindow(hwnd)
    if !IsObject(candidate)
        return
    ; Skip if already embedded or pending
    for host in GetAllHosts() {
        if host.tabRecords.Has(candidate.id)
            return
    }
    if State.PendingCandidates.Has(candidate.id)
        return
    ; Skip if content/top already embedded
    for host in GetAllHosts() {
        for tabId, record in host.tabRecords {
            if (record.contentHwnd = candidate.contentHwnd || record.topHwnd = candidate.topHwnd)
                return
        }
    }
    TryStackOrPending(State.MainHost, candidate)
}

; Shell hook: removes tab from pending or host when its window is destroyed.
OnWindowDestroyed(hwnd) {
    State.LastHookEventTick := A_TickCount
    ; Remove from pending if this hwnd was a candidate
    stalePending := []
    for tabId, pending in State.PendingCandidates {
        if (pending.candidate.topHwnd = hwnd)
            stalePending.Push(tabId)
    }
    for tabId in stalePending
        State.PendingCandidates.Delete(tabId)
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
                if State.MainHost && host = State.MainHost && Config.ShowOnlyWhenTabs && host.tabOrder.Length = 0
                    State.MainHost.gui.Hide()
                return
            }
        }
    }
}

; Slow-sweep timer: discovers new windows, updates titles, removes stale tabs, shows/hides host.
RefreshWindows(*) {

    if !State.MainHost || State.IsCleaningUp
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
                    UpdateRecordTitleCache(record)
                    titleChanged := true
                    tabIdOfLastTitleChange := tabId
                }
            }
        }

        ; Discovery only for main host (popouts don't scan for new windows)
        if host = State.MainHost {
            if TickElapsed(State.LastHookEventTick, now) >= 5000 {
                ; hooks have been quiet for 5s — run full discovery scan as fallback
                candidates := DiscoverCandidateWindows()
                for candidate in candidates {
                    currentIds[candidate.id] := true

                    if host.tabRecords.Has(candidate.id) {
                        if UpdateTrackedTab(host, candidate.id, candidate)
                            structureChanged := true
                        continue
                    }

                    if !State.PendingCandidates.Has(candidate.id) {
                        TryStackOrPending(host, candidate)
                        if host.tabRecords.Has(candidate.id)
                            structureChanged := true
                        continue
                    }

                    ; Already pending: refresh candidate metadata only, do NOT reset firstSeen
                    State.PendingCandidates[candidate.id].candidate := candidate
                }

                stalePending := []
                for tabId, pending in State.PendingCandidates {
                    if !currentIds.Has(tabId)
                        stalePending.Push(tabId)
                }
                for tabId in stalePending {
                    State.PendingCandidates.Delete(tabId)
                }
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
            if TickElapsed(record.lastSeenTick, now) > Config.TabDisappearGraceMs
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
    if Config.ShowOnlyWhenTabs && State.MainHost && IsWindowExists(State.MainHost.hwnd) {
        liveCount := 0
        for tabId in State.MainHost.tabOrder {
            if State.MainHost.tabRecords.Has(tabId) && IsWindowExists(State.MainHost.tabRecords[tabId].contentHwnd)
                liveCount++
        }
        if liveCount >= 1 {
            if !WinExist("ahk_id " State.MainHost.hwnd)
                State.MainHost.gui.Show("NoActivate")
        } else
            State.MainHost.gui.Hide()
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

    try {
        if !IsWindowExists(topHwnd)
            return ""
        title := WinGetTitle("ahk_id " topHwnd)
        if !DllCall("IsWindowVisible", "ptr", topHwnd)
            return ""
        if (title = "")
            return ""
        if Config.WindowTitleMatches.Length = 0
            return ""  ; No match patterns configured
        matched := false
        for pat in Config.WindowTitleMatches {
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
        if Config.TargetExe && (StrLower(processName) != StrLower(Config.TargetExe))
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
            hierarchySummary: Config.DebugDiscovery ? DescribeWindowHierarchy(topHwnd, contentHwnd) : ""
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
    for pat in Config.WindowTitleMatches {
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

; Builds a stable unique tab ID from PID, root owner, normalized title, and
; content window class. Does not include contentHwnd so the ID survives reflows.
BuildCandidateId(topHwnd, title, processName, contentHwnd) {
    pid := WinGetPID("ahk_id " topHwnd)
    rootOwner := GetRootOwner(topHwnd)
    contentClass := GetWindowClassName(contentHwnd)
    return pid "|" rootOwner "|" NormalizeTitle(title) "|" contentClass
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
        filteredTitle: FilterTitle(candidate.title),
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
    UpdateRecordTitleCache(record)
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
        for i, h in State.PopoutHosts {
            if h = host {
                State.PopoutHosts.RemoveAt(i)
                break
            }
        }
        if host.hwnd
            State.HostByHwnd.Delete(host.hwnd "")
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
    if !host.isPopout && Config.ShowOnlyWhenTabs && host.tabOrder.Length = 0
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
        record.sourceWasHidden := false
        AppendDebugLog("AttachTrackedWindow: SetParent failed for hwnd=" hwnd " tabId=" tabId "`r`n", true)
        SendMessage(0x000B, 1, 0,, "ahk_id " hwnd,,,, 500)  ; WM_SETREDRAW TRUE (re-enable drawing)
        return false
    }

    if (record.topHwnd != hwnd) && IsWindowExists(record.topHwnd) {
        record.sourceWasHidden := true
        DllCall("ShowWindow", "ptr", record.topHwnd, "int", SW_HIDE)
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
            windowWidth := Config.HostWidth
    }
    if !windowWidth
        windowWidth := Config.HostWidth

    if Config.TabPosition = "bottom" {
        if !windowHeight {
            windowHeight := GetClientHeight(host.hwnd)
            if !windowHeight
                windowHeight := Config.HostHeight
        }
        tabBarY := windowHeight - Config.HeaderHeight
    } else {
        tabBarY := Config.UseCustomTitleBar ? Config.TitleBarHeight : 0
    }
    if host.HasProp("titleBarBg") && host.titleBarBg {
        host.titleBarBg.Move(0, 0, windowWidth, Config.TitleBarHeight)
        host.titleText.Move(8, 0, windowWidth - 60, Config.TitleBarHeight)
        host.titleCloseBtn.Move(windowWidth - 46, 0, 46, Config.TitleBarHeight)
    }
    ; Resize tab bar background to full width
    if host.HasProp("tabBarBg") && host.tabBarBg
        host.tabBarBg.Move(0, tabBarY, windowWidth, Config.HeaderHeight)

    tabCount := host.tabOrder.Length

    if !tabCount {
        DrawTabBar(host)
        return
    }

    ; Compute tab Y: legacy TabBarOffsetY >= 0, or from alignment (top/center/bottom)
    if Config.TabBarOffsetY >= 0
        tabOffsetY := Config.TabBarOffsetY
    else {
        align := StrLower(Config.TabBarAlignment)
        if (align = "top")
            tabOffsetY := 0
        else if (align = "bottom")
            tabOffsetY := Config.HeaderHeight - Config.TabHeight
        else
            tabOffsetY := (Config.HeaderHeight - Config.TabHeight) // 2  ; center (default)
    }
    DrawTabBar(host)
    } finally {
    if host.hwnd && IsWindowExists(host.hwnd) {
        prev := DetectHiddenWindows(true)
        try SendMessage(0x000B, 1, 0,, "ahk_id " host.hwnd)  ; always re-enable
        DetectHiddenWindows(prev)
        try DllCall("RedrawWindow", "Ptr", host.hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x0085)
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
        State.PopoutHosts.Pop()
        if popoutHost.hwnd
            State.HostByHwnd.Delete(popoutHost.hwnd "")
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

    if !popoutHost.tabRecords.Has(tabId) || !popoutHost.isPopout
        return

    record := popoutHost.tabRecords[tabId]

    ; Add to main host
    State.MainHost.tabRecords[tabId] := record
    State.MainHost.tabOrder.Push(tabId)
    if (State.MainHost.activeTabId = "")
        State.MainHost.activeTabId := tabId

    ; Remove from popout
    for idx, currentId in popoutHost.tabOrder {
        if currentId = tabId {
            popoutHost.tabOrder.RemoveAt(idx)
            break
        }
    }
    popoutHost.tabRecords.Delete(tabId)

    if !TransferTrackedWindow(popoutHost, State.MainHost, tabId) {
        ; Failed - window never moved, restore data structures
        State.MainHost.tabRecords.Delete(tabId)
        for idx, currentId in State.MainHost.tabOrder {
            if currentId = tabId {
                State.MainHost.tabOrder.RemoveAt(idx)
                break
            }
        }
        if (State.MainHost.activeTabId = tabId)
            State.MainHost.activeTabId := State.MainHost.tabOrder.Length ? State.MainHost.tabOrder[1] : ""
        popoutHost.tabRecords[tabId] := record
        popoutHost.tabOrder.Push(tabId)
        popoutHost.activeTabId := tabId
        return
    }

    ; Destroy popout host
    for i, h in State.PopoutHosts {
        if h = popoutHost {
            State.PopoutHosts.RemoveAt(i)
            break
        }
    }
    if popoutHost.hwnd
        State.HostByHwnd.Delete(popoutHost.hwnd "")
    InvalidateHostsCache()
    popoutHost.gui.Destroy()

    LayoutTabButtons(State.MainHost)
    ShowOnlyActiveTab(State.MainHost)
    UpdateHostTitle(State.MainHost)
    RedrawAnyWindow(State.MainHost.hwnd)
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

    alignHVal := (Config.TabTitleAlignH = "left") ? 0 : (Config.TabTitleAlignH = "right") ? 2 : 1
    alignVVal := (Config.TabTitleAlignV = "top") ? 0 : (Config.TabTitleAlignV = "bottom") ? 2 : 1

    if !host.HasProp("tabCanvas") || !host.tabCanvas
        return
    if !host.hwnd || !IsWindowExists(host.hwnd)
        return

    w := GetClientWidth(host.hwnd)
    h := GetClientHeight(host.hwnd)
    if !w || !h
        return

    tabBarW := w
    tabBarH := Config.HeaderHeight
    tabBarY := Config.UseCustomTitleBar ? Config.TitleBarHeight : 0
    if Config.TabPosition = "bottom"
        tabBarY := h - Config.HeaderHeight

    if host.HasProp("tabBarBg") && host.tabBarBg
        host.tabBarBg.Move(0, tabBarY, tabBarW, tabBarH)
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
    DllCall("gdiplus\GdipCreateSolidFill", "UInt", HexToARGB(Config.ThemeTabBarBg),
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
    usableWidth := Max(200, tabBarW - (Config.HostPadding * 2))
    tabWidth := Floor((usableWidth - ((tabCount - 1) * Config.TabGap)) / tabCount)
    tabWidth := Max(Config.MinTabWidth, Min(Config.MaxTabWidth, tabWidth))
    ; Overflow: if tabs exceed available width, switch to scroll mode
    totalW := tabCount * tabWidth + (tabCount - 1) * Config.TabGap
    needScroll := totalW > usableWidth
    effectivePopoutW := Config.ShowPopoutButton ? Config.PopoutButtonWidth : 0
    effectiveCloseW  := Config.ShowCloseButton  ? Config.CloseButtonWidth  : 0
    titleWidth := tabWidth - effectivePopoutW - effectiveCloseW

    ; Clamp scroll offset and compute how many tabs fit
    if needScroll {
        visibleCount := Max(1, Floor((usableWidth - arrowW * 2) / (tabWidth + Config.TabGap)))
        host.tabScrollMax := Max(0, tabCount - visibleCount)
        host.tabScrollOffset := Max(0, Min(host.tabScrollOffset, host.tabScrollMax))
        drawStart := host.tabScrollOffset + 1   ; 1-based index into host.tabOrder
        drawEnd   := Min(tabCount, host.tabScrollOffset + visibleCount)
        x := Config.HostPadding + arrowW
    } else {
        host.tabScrollMax := 0
        host.tabScrollOffset := 0
        drawStart := 1
        drawEnd   := tabCount
        x := Config.HostPadding
    }

    if Config.TabBarOffsetY >= 0
        tabOffsetY := Config.TabBarOffsetY
    else {
        align := StrLower(Config.TabBarAlignment)
        if align = "top"
            tabOffsetY := 0
        else if align = "bottom"
            tabOffsetY := tabBarH - Config.TabHeight
        else
            tabOffsetY := (tabBarH - Config.TabHeight) // 2
    }

    ; Create three tab background brushes once (at most 3 distinct colors per redraw), reuse in loop, delete after
    pBrushActive := 0
    pBrushInactive := 0
    pBrushHover := 0
    DllCall("gdiplus\GdipCreateSolidFill", "UInt", HexToARGB(Config.ThemeTabActiveBg), "UPtr*", &pBrushActive)
    DllCall("gdiplus\GdipCreateSolidFill", "UInt", HexToARGB(Config.ThemeTabInactiveBg), "UPtr*", &pBrushInactive)
    DllCall("gdiplus\GdipCreateSolidFill", "UInt", HexToARGB(Config.ThemeTabInactiveBgHover), "UPtr*", &pBrushHover)

    for i, tabId in host.tabOrder {
        if needScroll && (i < drawStart || i > drawEnd)
            continue
        isActive  := (tabId = host.activeTabId)
        isHovered := (tabId = host.tabHoveredId)

        pTabBgBrush := isActive ? pBrushActive : (isHovered ? pBrushHover : pBrushInactive)
        fgColor := isActive  ? HexToARGB(Config.ThemeTabActiveText)
                 :             HexToARGB(Config.ThemeTabInactiveText)
        iconColor := isActive ? HexToARGB(Config.ThemeTabActiveText)
                   :            HexToARGB(Config.ThemeIconColor)

        ; Tab background with rounded corners (radius 5)
        GdipFillRoundRectWithBrush(pGraphics, x, tabOffsetY, tabWidth, Config.TabHeight, Config.TabCornerRadius, pTabBgBrush, HexToARGB(Config.ThemeTabBarBg))

        ; Active indicator strip
        if isActive && Config.TabIndicatorHeight > 0 {
            indicColor := HexToARGB(Config.ThemeTabIndicatorColor != ""
                ? Config.ThemeTabIndicatorColor : Config.ThemeTabActiveBg)
            indicY := (Config.TabPosition = "bottom")
                ? tabOffsetY
                : tabOffsetY + Config.TabHeight - Config.TabIndicatorHeight
            indicW := Max(1, tabWidth - 8)
            ; Radius must not exceed half the height — GDI+ arc math goes negative otherwise.
            ; Use 0 (sharp rect) when indicator is thinner than 3px.
            indicR := (Config.TabIndicatorHeight >= 3) ? 1 : 0
            GdipFillRoundRect(pGraphics, x + 4, indicY, indicW, Config.TabIndicatorHeight, indicR, indicColor)
        }

        ; Tab title
        rawTitle := host.tabRecords.Has(tabId)
            ? host.tabRecords[tabId].filteredTitle : "Window"
        if Config.ShowTabNumbers
            rawTitle := i ". " rawTitle
        GdipDrawStringSimple(pGraphics, rawTitle,
            x, tabOffsetY, titleWidth, Config.TabHeight,
            fgColor, Config.ThemeFontNameTab, Config.ThemeFontSize, isActive,
            true, true, alignHVal, alignVVal)

        ; Popout / merge icon (only drawn if ShowPopoutButton)
        if Config.ShowPopoutButton {
            iconText := host.isPopout ? Config.IconMerge : Config.IconPopout
            GdipDrawStringSimple(pGraphics, iconText,
                x + titleWidth, tabOffsetY, Config.PopoutButtonWidth, Config.TabHeight,
                iconColor, Config.ThemeIconFont, Config.ThemeIconFontSize, false)
        }

        ; Close icon (only drawn if ShowCloseButton)
        if Config.ShowCloseButton {
            GdipDrawStringSimple(pGraphics, Config.IconClose,
                x + titleWidth + Config.PopoutButtonWidth, tabOffsetY,
                Config.CloseButtonWidth, Config.TabHeight,
                iconColor, Config.ThemeIconFont, Config.ThemeIconFontSize, false)
        }

        x += tabWidth + Config.TabGap
    }

    DllCall("gdiplus\GdipDeleteBrush", "UPtr", pBrushActive)
    DllCall("gdiplus\GdipDeleteBrush", "UPtr", pBrushInactive)
    DllCall("gdiplus\GdipDeleteBrush", "UPtr", pBrushHover)

    ; Vertical separators between tabs
    if Config.TabSeparatorWidth > 0 {
        sepColor := Config.ThemeTabSeparatorColor != "" ? HexToARGB(Config.ThemeTabSeparatorColor) : HexToARGB(Config.ThemeContentBorder)
        pSepBrush := 0
        DllCall("gdiplus\GdipCreateSolidFill", "UInt", sepColor, "UPtr*", &pSepBrush)
        sepX := needScroll ? Config.HostPadding + arrowW : Config.HostPadding
        for i, tabId in host.tabOrder {
            if needScroll && (i < drawStart || i > drawEnd)
                continue
            if (needScroll ? i < drawEnd : i < tabCount) {
                DllCall("gdiplus\GdipFillRectangleI", "UPtr", pGraphics, "UPtr", pSepBrush,
                    "Int", sepX + tabWidth, "Int", tabOffsetY, "Int", Config.TabSeparatorWidth, "Int", Config.TabHeight)
            }
            sepX += tabWidth + Config.TabGap
        }
        DllCall("gdiplus\GdipDeleteBrush", "UPtr", pSepBrush)
    }

    ; Draw left scroll arrow only when can scroll left (tabScrollOffset > 0)
    if needScroll && host.tabScrollOffset > 0 {
        arrowColor := HexToARGB(Config.ThemeTabActiveText)
        GdipDrawStringSimple(pGraphics, Chr(0xE76B),
            Config.HostPadding, tabOffsetY, arrowW, Config.TabHeight,
            arrowColor, Config.ThemeIconFont, Config.ThemeIconFontSize, false)
    }

    ; Draw right scroll arrow only when can scroll right (tabScrollOffset < tabScrollMax)
    if needScroll && host.tabScrollOffset < host.tabScrollMax {
        arrowColor := HexToARGB(Config.ThemeTabActiveText)
        GdipDrawStringSimple(pGraphics, Chr(0xE76C),
            Config.HostPadding + arrowW + (drawEnd - drawStart + 1) * (tabWidth + Config.TabGap) - Config.TabGap,
            tabOffsetY, arrowW, Config.TabHeight,
            arrowColor, Config.ThemeIconFont, Config.ThemeIconFontSize, false)
    }

    ApplyBitmapToCanvas(host.tabCanvas, pBitmap, pGraphics)
}

; Updates host window title with tab count and active tab name.
UpdateHostTitle(host) {

    if !host || !host.gui
        return

    liveCount := GetLiveTabCount(host)
    if host.isPopout
        suffix := " (popped out)"
    else
        suffix := ""
    if (host.activeTabId != "") && host.tabRecords.Has(host.activeTabId)
        title := Config.HostTitle . " (" liveCount ") - " . host.tabRecords[host.activeTabId].title . suffix
    else
        title := Config.HostTitle . " (" liveCount ")" . suffix
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

    padBottom := (Config.HostPaddingBottom >= 0) ? Config.HostPaddingBottom : Config.HostPadding
    x := Config.HostPadding
    customTitleH := Config.UseCustomTitleBar ? Config.TitleBarHeight : 0

    if Config.TabPosition = "bottom"
        y := customTitleH + Config.HostPadding
    else
        y := customTitleH + Config.HeaderHeight + Config.HostPadding

    if host.hwnd && IsWindowExists(host.hwnd) {
        try {
            WinGetClientPos(,, &clientW, &clientH, "ahk_id " host.hwnd)
            w := Max(200, clientW - (Config.HostPadding * 2))
            if Config.TabPosition = "bottom"
                h := Max(140, clientH - y - Config.HeaderHeight - padBottom)
            else
                h := Max(140, clientH - y - padBottom)
            return
        }
    }

    w := Max(200, Config.HostWidth - (Config.HostPadding * 2))
    if Config.TabPosition = "bottom"
        h := Max(140, Config.HostHeight - y - Config.HeaderHeight - padBottom)
    else
        h := Max(140, Config.HostHeight - y - padBottom)
}

; Exit handler: restores all tabs, clears state.
CleanupAll(*) {

    if State.IsCleaningUp
        return

    State.IsCleaningUp := true
    State.WatchdogTimerActive := false
    SetTimer(WatchdogCheck, 0)
    SetTimer(TooltipCheckTimer, 0)

    if State.WinEventHooks.Length {
        for hook in State.WinEventHooks
            DllCall("UnhookWinEvent", "Ptr", hook)
        State.WinEventHooks := []
    }

    for host in GetAllHosts() {
        for tabId in host.tabOrder.Clone()
            RemoveTrackedTab(host, tabId, true)
    }

    ; Destroy any remaining icon handles on all hosts
    for host in GetAllHosts() {
        if host.HasProp("iconHandle") && host.iconHandle {
            DllCall("DestroyIcon", "ptr", host.iconHandle)
            host.iconHandle := 0
        }
    }

    State.PendingCandidates := Map()
    ; Release cached GDI+ font objects
    for _, pFamily in State.CachedFontFamily
        DllCall("gdiplus\GdipDeleteFontFamily", "UPtr", pFamily)
    State.CachedFontFamily := Map()
    for _, pFont in State.CachedFont
        DllCall("gdiplus\GdipDeleteFont", "UPtr", pFont)
    State.CachedFont := Map()
    if IsObject(State.CachedStringFormat) {
        for _, pFmt in State.CachedStringFormat
            DllCall("gdiplus\GdipDeleteStringFormat", "UPtr", pFmt)
        State.CachedStringFormat := 0
    }
    if State.GdipToken {
        DllCall("gdiplus\GdiplusShutdown", "UPtr", State.GdipToken)
        State.GdipToken := 0
    }
    State.IsCleaningUp := false
}

; Writes discovered candidate windows to discovery.txt (when DebugDiscovery=1).
DumpDiscoveryDebug() {

    discovered := DiscoverCandidateWindows()
    text := "Timestamp: " FormatTime(, "yyyy-MM-dd HH:mm:ss") "`r`n"
    text .= "Discovered windows: " discovered.Length "`r`n`r`n"

    for candidate in discovered {
        text .= "Tab ID: " candidate.id "`r`n"
        text .= candidate.hierarchySummary "`r`n"
        text .= "------------------------------`r`n"
    }

    FileDelete(Config.DebugLogPath)
    FileAppend(text, Config.DebugLogPath, "UTF-8")
    MsgBox("Wrote discovery info to:`n" Config.DebugLogPath, "StackTabs Debug")
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
    hBadged := CreateBadgedIcon(hSource, Config.ThemeTabActiveBg)
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
    if !Config.DebugDiscovery && !critical
        return
    FileAppend("[" FormatTime(, "yyyy-MM-dd HH:mm:ss") "]`r`n" text, Config.DebugLogPath, "UTF-8")
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
    for pattern in Config.TitleStripPatterns
        title := RegExReplace(title, pattern, "")
    return Trim(title)
}

; Updates cached filtered title on a record. Call whenever record.title changes.
UpdateRecordTitleCache(record) {
    record.filteredTitle := FilterTitle(record.title)
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
    tabCount := host.tabOrder.Length
    if tabCount = 0
        return 0
    w := GetClientWidth(host.hwnd)
    usableWidth := Max(200, w - (Config.HostPadding * 2))
    tabWidth := Floor((usableWidth - ((tabCount - 1) * Config.TabGap)) / tabCount)
    tabWidth := Max(Config.MinTabWidth, Min(Config.MaxTabWidth, tabWidth))
    return tabWidth
}

GetTabIndexAtMouseX(host, mouseX) {
    tabWidth := GetTabWidthForHost(host)
    if tabWidth = 0
        return 0
    tabCount := host.tabOrder.Length
    arrowW := 24
    needScroll := host.HasProp("tabScrollMax") && host.tabScrollMax > 0
    startX    := needScroll ? Config.HostPadding + arrowW : Config.HostPadding
    startIdx  := needScroll ? host.tabScrollOffset + 1 : 1
    Loop tabCount {
        i := A_Index
        if i < startIdx
            continue
        slotX := startX + (i - startIdx) * (tabWidth + Config.TabGap)
        if mouseX >= slotX && mouseX < slotX + tabWidth
            return i
    }
    return 0
}

GetTabZoneAtMouseX(host, mouseX, tabIdx) {
    tabWidth := GetTabWidthForHost(host)
    arrowW := 24
    needScroll := host.HasProp("tabScrollMax") && host.tabScrollMax > 0
    startX   := needScroll ? Config.HostPadding + arrowW : Config.HostPadding
    startIdx := needScroll ? host.tabScrollOffset + 1 : 1
    slotX := startX + (tabIdx - startIdx) * (tabWidth + Config.TabGap)
    effectivePopoutW := Config.ShowPopoutButton ? Config.PopoutButtonWidth : 0
    effectiveCloseW  := Config.ShowCloseButton  ? Config.CloseButtonWidth  : 0
    titleWidth := tabWidth - effectiveCloseW - effectivePopoutW
    if mouseX < slotX + titleWidth
        return "title"
    if mouseX < slotX + titleWidth + Config.PopoutButtonWidth
        return "popout"
    return "close"
}

; Returns true if the cursor is over any host's tab bar (screen coords).
IsMouseOverAnyTabBar() {
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
        tabBarY := Config.UseCustomTitleBar ? Config.TitleBarHeight : 0
        if Config.TabPosition = "bottom"
            tabBarY := clientH - Config.HeaderHeight
        if clientY >= tabBarY && clientY < tabBarY + Config.HeaderHeight
            return true
    }
    return false
}

OnTabCanvasMouseMove(wParam, lParam, msg, hwnd) {
    mouseX := lParam & 0xFFFF
    mouseY := (lParam >> 16) & 0xFFFF
    for host in GetAllHosts() {
        if host.hwnd != hwnd && (!host.HasProp("tabCanvas") || host.tabCanvas.Hwnd != hwnd)
            continue
        tabBarY := Config.UseCustomTitleBar ? Config.TitleBarHeight : 0
        if Config.TabPosition = "bottom" {
            clientH := GetClientHeight(host.hwnd)
            tabBarY := clientH - Config.HeaderHeight
        }
        checkY := (hwnd = host.tabCanvas.Hwnd) ? 0 : tabBarY
        newHoveredId := ""
        if mouseY >= checkY && mouseY < checkY + Config.HeaderHeight {
            tabIdx := GetTabIndexAtMouseX(host, mouseX)
            newHoveredId := (tabIdx > 0) ? host.tabOrder[tabIdx] : ""
        }
        if newHoveredId != host.tabHoveredId {
            host.tabHoveredId := newHoveredId
            DrawTabBar(host)
        }
        if newHoveredId != "" && host.tabRecords.Has(newHoveredId) {
            fullTitle := host.tabRecords[newHoveredId].filteredTitle
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
    arrowW := 24
    mouseX := lParam & 0xFFFF
    mouseY := (lParam >> 16) & 0xFFFF
    for host in GetAllHosts() {
        if host.hwnd != hwnd && (!host.HasProp("tabCanvas") || host.tabCanvas.Hwnd != hwnd)
            continue
        tabBarY := Config.UseCustomTitleBar ? Config.TitleBarHeight : 0
        if Config.TabPosition = "bottom" {
            clientH := GetClientHeight(host.hwnd)
            tabBarY := clientH - Config.HeaderHeight
        }
        ; When message is from tab canvas, coords are canvas-relative (tab bar origin = 0,0)
        checkY := (hwnd = host.tabCanvas.Hwnd) ? 0 : tabBarY
        if mouseY < checkY || mouseY >= checkY + Config.HeaderHeight
            continue
        ; Scroll arrows — only handle when arrow is visible (can scroll in that direction)
        if host.tabScrollMax > 0 {
            if host.tabScrollOffset > 0 && mouseX >= Config.HostPadding && mouseX < Config.HostPadding + arrowW {
                host.tabScrollOffset := Max(0, host.tabScrollOffset - 1)
                DrawTabBar(host)
                return
            }
            tabWidth := GetTabWidthForHost(host)
            visibleCount := Max(1, Floor((GetClientWidth(host.hwnd) - (Config.HostPadding * 2) - arrowW * 2) / (tabWidth + Config.TabGap)))
            rightArrowX := Config.HostPadding + arrowW + visibleCount * (tabWidth + Config.TabGap) - Config.TabGap
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
    mouseX := lParam & 0xFFFF
    mouseY := (lParam >> 16) & 0xFFFF
    for host in GetAllHosts() {
        if host.hwnd != hwnd && (!host.HasProp("tabCanvas") || host.tabCanvas.Hwnd != hwnd)
            continue
        tabBarY := Config.UseCustomTitleBar ? Config.TitleBarHeight : 0
        if Config.TabPosition = "bottom" {
            clientH := GetClientHeight(host.hwnd)
            tabBarY := clientH - Config.HeaderHeight
        }
        checkY := (hwnd = host.tabCanvas.Hwnd) ? 0 : tabBarY
        if mouseY < checkY || mouseY >= checkY + Config.HeaderHeight
            return
        tabIdx := GetTabIndexAtMouseX(host, mouseX)
        if tabIdx > 0 && tabIdx <= host.tabOrder.Length
            CloseTab(host, host.tabOrder[tabIdx])
        return
    }
}

OnTabCanvasRightClick(wParam, lParam, msg, hwnd) {
    mouseX := lParam & 0xFFFF
    mouseY := (lParam >> 16) & 0xFFFF
    for host in GetAllHosts() {
        if host.hwnd != hwnd && (!host.HasProp("tabCanvas") || host.tabCanvas.Hwnd != hwnd)
            continue
        tabBarY := Config.UseCustomTitleBar ? Config.TitleBarHeight : 0
        if Config.TabPosition = "bottom" {
            clientH := GetClientHeight(host.hwnd)
            tabBarY := clientH - Config.HeaderHeight
        }
        checkY := (hwnd = host.tabCanvas.Hwnd) ? 0 : tabBarY
        if mouseY < checkY || mouseY >= checkY + Config.HeaderHeight
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

; Safe elapsed time that handles A_TickCount wraparound (~49.7 day boundary).
TickElapsed(start, now) => (now >= start) ? (now - start) : (0xFFFFFFFF - start + now)

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

; Fills a rounded rect using an existing brush (caller owns the brush). Used to reuse brushes across tab draws.
; bgArgbColor clears corner pixels before drawing the rounded shape (avoids transparent corners when composited).
GdipFillRoundRectWithBrush(pGraphics, x, y, w, h, radius, pBrush, bgArgbColor) {
    if radius <= 0 {
        DllCall("gdiplus\GdipFillRectangleI", "UPtr", pGraphics, "UPtr", pBrush,
            "Int", x, "Int", y, "Int", w, "Int", h)
        return
    }
    ; Fill full bounding rect with bg color first to clear corners
    pBgBrush := 0
    DllCall("gdiplus\GdipCreateSolidFill", "UInt", bgArgbColor, "UPtr*", &pBgBrush)
    DllCall("gdiplus\GdipFillRectangleI", "UPtr", pGraphics, "UPtr", pBgBrush,
        "Int", x, "Int", y, "Int", w, "Int", h)
    DllCall("gdiplus\GdipDeleteBrush", "UPtr", pBgBrush)
    ; Draw rounded shape on top
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
    DllCall("gdiplus\GdipFillPath", "UPtr", pGraphics,
        "UPtr", pBrush, "UPtr", pPath)
    DllCall("gdiplus\GdipDeletePath", "UPtr", pPath)
}

GdipDrawStringSimple(pGraphics, text, x, y, w, h, argbColor, fontFamilyName, fontSize, bold, noWrap := true, ellipsis := true, alignH := 1, alignV := 1) {

    ; Build cache keys
    familyKey := fontFamilyName
    fontKey := fontFamilyName "|" fontSize "|" (bold ? 1 : 0)

    ; Create or reuse font family
    if !State.CachedFontFamily.Has(familyKey) {
        pFamily := 0
        DllCall("gdiplus\GdipCreateFontFamilyFromName",
            "Str", fontFamilyName, "Ptr", 0, "UPtr*", &pFamily)
        if !pFamily
            return
        State.CachedFontFamily[familyKey] := pFamily
    }
    pFamily := State.CachedFontFamily[familyKey]

    ; Create or reuse font
    if !State.CachedFont.Has(fontKey) {
        pFont := 0
        DllCall("gdiplus\GdipCreateFont",
            "UPtr", pFamily, "Float", Float(fontSize),
            "Int", bold ? 1 : 0, "Int", 3, "UPtr*", &pFont)
        if !pFont
            return
        State.CachedFont[fontKey] := pFont
    }
    pFont := State.CachedFont[fontKey]

    ; GDI+ StringAlignment: 0=Near/left, 1=Center, 2=Far/right
    fmtKey := (noWrap ? "1" : "0") "|" (ellipsis ? "1" : "0") "|" alignH "|" alignV
    if !State.CachedStringFormat || !State.CachedStringFormat.Has(fmtKey) {
        if !IsObject(State.CachedStringFormat)
            State.CachedStringFormat := Map()
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
        State.CachedStringFormat[fmtKey] := pFormat
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
        "UPtr", pFont, "Ptr", rect, "UPtr", State.CachedStringFormat[fmtKey], "UPtr", pBrush)
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
        tabBarY := Config.UseCustomTitleBar ? Config.TitleBarHeight : 0
        if Config.TabPosition = "bottom" {
            clientH := GetClientHeight(host.hwnd)
            tabBarY := clientH - Config.HeaderHeight
        }
        if clientY < tabBarY || clientY >= tabBarY + Config.HeaderHeight
            continue
        if host.tabScrollMax > 0 {
            host.tabScrollOffset := Max(0, Min(host.tabScrollMax, host.tabScrollOffset - delta))
            DrawTabBar(host)
        }
        return
    }
}
OnMessage(0x020A, OnTabCanvasMouseWheel)

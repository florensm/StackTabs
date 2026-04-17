; StackTabs - owner-aware embedded window host
; AutoHotkey v2

#Requires AutoHotkey v2.0
#SingleInstance Force

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
    KeepAboveTabApps: false,   ; Keep host above any window in a tracked tab's process whenever that process takes foreground, and reparent its dialogs to the host so they stay above it via the OS owner rule
    KeepAboveTabAppsDebug: false,  ; Verbose z-order trace to debug-zorder.log (paired with KeepAboveTabApps)

    ; === THEME (loaded from themes\ folder; dark.ini is the default and fallback) ===
    ThemeTabIndicatorColor: "",   ; set by LoadThemeFromFile; defaults to TabActiveBg
    ThemeIconFont: "",   ; auto-detected at startup; override in theme file with IconFont=
    ActiveThemeFile: "dark.ini",   ; overridden by ThemeFile= in config.ini
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
    SwitcherCards: [],
    SwitcherCtrlTabMode: false,
    SwitcherOrigTabIdx: 0,
    ZOrderEnforcerFn: "",      ; bound EnforceZOrderTick closure; "" = not running
    ZOrderEnforceBusy: false,  ; reentrancy guard for EnforceHostZOrder
    LastActiveTrackedPopup: ""  ; { host, hwnd } of the last tracked owned window that was foreground; used to redirect focus to the host when the popup closes
}

; Returns the scalar value of an INI entry with any inline "; ..." comment
; stripped and surrounding whitespace trimmed. Windows INI format treats ';'
; only as a line comment, not an end-of-value comment, so IniRead returns
; the full string including the comment text. Raw comparisons and Integer()
; conversions on such strings fail silently (the failure is caught by the
; outer try in LoadConfigFromIni, skipping everything after the first throw).
IniClean(path, section, key, default := "") {
    raw := IniRead(path, section, key, "")
    if raw = ""
        return default
    parts := StrSplit(raw, ";")
    return Trim(parts[1])
}

IniBool(path, section, key, default := false) {
    val := IniClean(path, section, key, "")
    if val = ""
        return default
    return val = "1"
}

IniInt(path, section, key, default) {
    val := IniClean(path, section, key, "")
    if val = ""
        return default
    try return Integer(val)
    return default
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
        Config.TargetExe := IniClean(iniPath, "General", "TargetExe", Config.TargetExe)
        Config.SlowSweepInterval := IniInt(iniPath, "General", "SlowSweepInterval", Config.SlowSweepInterval)
        Config.StackDelayMs := IniInt(iniPath, "General", "StackDelayMs", Config.StackDelayMs)
        Config.StackSwitchDelayMs := IniInt(iniPath, "General", "StackSwitchDelayMs", Config.StackSwitchDelayMs)
        Config.WatchdogMaxMs := IniInt(iniPath, "General", "WatchdogMaxMs", Config.WatchdogMaxMs)
        Config.TabDisappearGraceMs := IniInt(iniPath, "General", "TabDisappearGraceMs", Config.TabDisappearGraceMs)
        Config.DebugDiscovery := IniBool(iniPath, "General", "DebugDiscovery", false)
        Config.KeepAboveTabApps := IniBool(iniPath, "General", "KeepAboveTabApps", false)
        Config.KeepAboveTabAppsDebug := IniBool(iniPath, "General", "KeepAboveTabAppsDebug", false)
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
    A_TrayMenu.Add("Reload", (*) => Reload())
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
; EVENT_SYSTEM_FOREGROUND (0x0003): a new top-level window becomes the foreground window.
; Used by the KeepAboveTabApps feature to bump the host above any activated tab-app window.
State.WinEventHooks.Push(DllCall("SetWinEventHook",
    "UInt", 0x0003, "UInt", 0x0003,
    "Ptr",  0,      "Ptr",  State.WinEventHookCallback,
    "UInt", 0,      "UInt", 0, "UInt", 0, "Ptr"))

; Persistent z-order enforcer for KeepAboveTabApps. Event-driven foreground
; handling alone cannot cover apps that raise their main shell via
; BringWindowToTop / explicit SetWindowPos without producing a foreground
; change. A low-frequency poll walks the z-order and nudges the host back
; into place whenever the invariant is violated.
if Config.KeepAboveTabApps
    StartZOrderEnforcer()
InitZOrderDebugLog()

OnExit(CleanupAll)

RefreshWindows()
; Slow sweep fallback: only does full WinGetList scan when hooks have been
; quiet for 5+ seconds. Primary discovery is via shell hook and WinEvent hooks.
SetTimer(RefreshWindows, Config.SlowSweepInterval)

; ============ TAB SWITCHER OVERLAY ============
; Ctrl+Tab: open switcher and cycle forward; Ctrl+Shift+Tab cycles backward.
; Keep Ctrl held to continue cycling; release Ctrl to commit. Escape cancels.

State.SwitcherGui         := ""
State.SwitcherAllTabs     := []
State.SwitcherVisible     := []
State.SwitcherSelVisIdx   := 0
State.SwitcherCards       := []
State.SwitcherCtrlTabMode := false   ; true = opened via Ctrl+Tab (no search box, Ctrl-release commits)
State.SwitcherOrigTabIdx  := 0       ; all-tabs index active when switcher opened (for Escape restore)

; Ctrl+Tab / Ctrl+Shift+Tab when StackTabs (or embedded content) is active
#HotIf StackTabsHostIsActive()
^Tab::  SwitcherCtrlTabCycle(1)
^+Tab:: SwitcherCtrlTabCycle(-1)
#HotIf

; Release Ctrl to commit while switcher is open (OnMessage WM_KEYUP is unreliable for
; modifier keys — use AHK hotkeys instead so the hook fires before the GUI sees it)
#HotIf State.SwitcherCtrlTabMode
LControl Up:: SwitcherCtrlRelease()
RControl Up:: SwitcherCtrlRelease()
#HotIf

SwitcherCtrlTabCycle(dir) {
    if State.SwitcherGui && State.SwitcherCtrlTabMode {
        SwitcherCycleStep(dir)
        return
    }
    ShowTabSwitcher(true, dir)
}

SwitcherCycleStep(dir) {
    count := State.SwitcherVisible.Length
    if count = 0
        return
    n := Mod(State.SwitcherSelVisIdx - 1 + dir + count * 2, count)
    State.SwitcherSelVisIdx := n + 1
    SwitcherRefreshCards()
    SwitcherPreviewSelected()
}

; Bring the tab-switcher GUI to the foreground if it is still a valid window.
SwitcherFocusGui() {
    if !State.SwitcherGui
        return
    try hid := State.SwitcherGui.Hwnd
    catch {
        State.SwitcherGui := ""
        return
    }
    if !hid || !IsWindowExists(hid)
        return
    prev := DetectHiddenWindows(true)
    try WinActivate("ahk_id " hid)
    DetectHiddenWindows(prev)
}

; Switch tab content + update tab bar; in Ctrl+Tab mode re-activate overlay (WinEvent may focus embedded app).
SwitcherPreviewSelected() {
    if !State.SwitcherGui
        return
    count := State.SwitcherVisible.Length
    if State.SwitcherSelVisIdx < 1 || State.SwitcherSelVisIdx > count
        return
    tabIdx := State.SwitcherVisible[State.SwitcherSelVisIdx]
    if tabIdx < 1 || tabIdx > State.SwitcherAllTabs.Length
        return
    item := State.SwitcherAllTabs[tabIdx]
    if !item.host.tabRecords.Has(item.tabId)
        return  ; tab was closed while switcher was open
    record := item.host.tabRecords[item.tabId]
    contentHwnd := record.contentHwnd
    if !IsWindowExists(contentHwnd)
        return
    item.host.activeTabId := item.tabId
    ; Skip the deferred WPF-settle repaint during Ctrl-held preview: it schedules a
    ; 50ms+20ms timer chain that piles up on rapid cycling. The final commit via
    ; SelectTab (on Ctrl release or Enter) runs ShowOnlyActiveTab with deferred=true
    ; so the settle pass still happens exactly once when the user lands on a tab.
    ShowOnlyActiveTab(item.host, !State.SwitcherCtrlTabMode)
    UpdateHostTitle(item.host)
    ; WinEvent NAMECHANGE may call SelectTab -> FocusEmbeddedContent while previewing; keep overlay focused for Tab/J-K.
    if State.SwitcherCtrlTabMode
        SwitcherFocusGui()
}

ShowTabSwitcher(ctrlTabMode := false, dir := 1) {
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
    if ctrlTabMode && allTabs.Length = 1
        return  ; nothing to cycle — skip overlay

    State.SwitcherCtrlTabMode := ctrlTabMode
    State.SwitcherAllTabs     := allTabs
    State.SwitcherVisible     := []
    State.SwitcherSelVisIdx   := 1
    Loop allTabs.Length
        State.SwitcherVisible.Push(A_Index)

    ; Find the currently active tab index
    origIdx := 1
    for idx, item in allTabs {
        if item.isActive {
            origIdx := idx
            break
        }
    }
    State.SwitcherOrigTabIdx := origIdx

    if ctrlTabMode {
        ; Start selection on the next/prev tab relative to the active one
        count := allTabs.Length
        State.SwitcherSelVisIdx := Mod(origIdx - 1 + dir + count * 2, count) + 1
    } else {
        State.SwitcherSelVisIdx := origIdx
    }

    ; Layout — vertical list (command-palette style)
    overlayWidth    := 420
    rowHeight       := 36
    pad             := 16
    listAreaHeight  := Min(allTabs.Length * rowHeight, 600)

    if ctrlTabMode {
        listAreaY     := pad
        overlayHeight := pad + listAreaHeight + pad
    } else {
        searchBarHeight := 38
        listGap       := 8
        listAreaY     := pad + searchBarHeight + listGap
        overlayHeight := pad + searchBarHeight + listGap + listAreaHeight + pad
    }

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

    if !ctrlTabMode {
        ; Search box — cue-banner text via EM_SETCUEBANNER after show
        State.SwitcherGui.SetFont("s" (Config.ThemeFontSize + 1) " c" Config.ThemeWindowText, Config.ThemeFontName)
        searchBox := State.SwitcherGui.Add("Edit",
            "x" pad " y" pad " w" (overlayWidth - pad * 2) " h" searchBarHeight
            " Background" Config.ThemeTabBarBg " c" Config.ThemeWindowText, "")
        searchBox.SetFont("s" (Config.ThemeFontSize + 1) " c" Config.ThemeWindowText, Config.ThemeFontName)
    }

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
    if !IsObject(State.SwitcherGui)
        return
    switcherHwnd := State.SwitcherGui.Hwnd
    DllCall("dwmapi.dll\DwmSetWindowAttribute",
        "ptr", switcherHwnd, "uint", 33, "uint*", 2, "uint", 4)

    OnMessage(0x0100, SwitcherOnKeyDown, 1)
    if ctrlTabMode {
        SwitcherPreviewSelected()   ; preview + SwitcherFocusGui when Ctrl+Tab mode
    } else {
        ; Cue banner "Search tabs..." inside the edit box
        SendMessage(0x1501, 1, StrPtr("Search tabs..."),, "ahk_id " searchBox.Hwnd)
        searchBox.OnEvent("Change", SwitcherOnSearch)
        SwitcherFocusGui()
        searchBox.Focus()
    }
}

SwitcherClose() {
    OnMessage(0x0100, SwitcherOnKeyDown, 0)
    State.SwitcherCtrlTabMode := false
    State.SwitcherOrigTabIdx  := 0
    if State.SwitcherGui {
        State.SwitcherGui.Destroy()
        State.SwitcherGui := ""
    }
    State.SwitcherAllTabs := []
    State.SwitcherVisible := []
    State.SwitcherCards   := []
    ; The overlay (AlwaysOnTop) covered the host while DrawTabBar ran during preview, so
    ; UpdateWindow(canvas) validated the canvas update region while DWM still showed the old
    ; frame. Now that the overlay is gone, force a fresh redraw so the correct bitmap shows.
    for host in GetAllHosts() {
        DrawTabBar(host)
        RedrawAnyWindow(host.hwnd)
    }
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
    ; Invisible cards can't receive clicks, so no visibility check needed
    SwitcherActivate(ctrl.tabSwitcherIdx)
}

SwitcherActivate(tabIdx) {
    if tabIdx < 1 || tabIdx > State.SwitcherAllTabs.Length
        return
    item := State.SwitcherAllTabs[tabIdx]
    ; Activate the host BEFORE destroying the overlay. The overlay is +AlwaysOnTop so it
    ; stays visible while the host renders underneath — when the overlay is destroyed last
    ; the host is already fully painted, eliminating the flash on close.
    if item.host.activeTabId = item.tabId && item.host.tabRecords.Has(item.tabId) {
        contentHwnd := item.host.tabRecords[item.tabId].contentHwnd
        if IsWindowExists(contentHwnd)
            FocusEmbeddedContent(item.host.hwnd, contentHwnd)
        else
            SelectTab(item.host, item.tabId)
    } else {
        SelectTab(item.host, item.tabId)
    }
    if item.host.hwnd && IsWindowExists(item.host.hwnd)
        WinActivate("ahk_id " item.host.hwnd)
    SwitcherClose()
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
    if wParam = 0x1B {  ; Escape
        if State.SwitcherCtrlTabMode && State.SwitcherOrigTabIdx >= 1
                && State.SwitcherOrigTabIdx <= State.SwitcherAllTabs.Length {
            ; Restore the original tab — activate host first, destroy overlay last (no flash)
            item := State.SwitcherAllTabs[State.SwitcherOrigTabIdx]
            SelectTab(item.host, item.tabId)
            if item.host.hwnd && IsWindowExists(item.host.hwnd)
                WinActivate("ahk_id " item.host.hwnd)
            SwitcherClose()
        } else {
            SwitcherClose()
        }
        return 0
    }
    if count = 0
        return
    if wParam = 0x09 && State.SwitcherCtrlTabMode {  ; Tab key — cycle while Ctrl held
        SwitcherCycleStep(GetKeyState("Shift") ? -1 : 1)
        return 0
    }
    if wParam = 0x0D {  ; Enter
        if State.SwitcherSelVisIdx >= 1 && State.SwitcherSelVisIdx <= count
            SwitcherActivate(State.SwitcherVisible[State.SwitcherSelVisIdx])
        return 0
    }
    if wParam = 0x25 || wParam = 0x26 {  ; Left / Up
        State.SwitcherSelVisIdx := Max(1, State.SwitcherSelVisIdx - 1)
        SwitcherRefreshCards()
        SwitcherPreviewSelected()
        return 0
    }
    if wParam = 0x27 || wParam = 0x28 {  ; Right / Down
        State.SwitcherSelVisIdx := Min(count, State.SwitcherSelVisIdx + 1)
        SwitcherRefreshCards()
        SwitcherPreviewSelected()
        return 0
    }
    if State.SwitcherCtrlTabMode {  ; J/K vim keys — only in ctrl-tab mode (no search box)
        if wParam = 0x4B {  ; K — up (wraps)
            State.SwitcherSelVisIdx := Mod(State.SwitcherSelVisIdx - 2 + count, count) + 1
            SwitcherRefreshCards()
            SwitcherPreviewSelected()
            return 0
        }
        if wParam = 0x4A {  ; J — down (wraps)
            State.SwitcherSelVisIdx := Mod(State.SwitcherSelVisIdx, count) + 1
            SwitcherRefreshCards()
            SwitcherPreviewSelected()
            return 0
        }
    }
}

SwitcherCtrlRelease() {
    if !State.SwitcherGui || !State.SwitcherCtrlTabMode
        return
    count := State.SwitcherVisible.Length
    if State.SwitcherSelVisIdx >= 1 && State.SwitcherSelVisIdx <= count
        SwitcherActivate(State.SwitcherVisible[State.SwitcherSelVisIdx])
    else
        SwitcherClose()
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
        ShowOnlyActiveTab(State.MainHost)
    }
}

; Win+Shift+D: dump discovery debug to disk (only when DebugDiscovery=1).
#+d:: {
    if !Config.DebugDiscovery
        return
    DumpDiscoveryDebug()
}

#HotIf StackTabsHostIsActive()
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

; WM_ACTIVATE handler: when StackTabs regains focus from an external popup (e.g. espanso's
; selection menu), redirect keyboard focus to the active embedded tab so input goes there.
OnWmActivate(wParam, lParam, msg, hwnd) {
    if (wParam & 0xFFFF) = 0  ; being deactivated, not activated
        return
    ; Don't fight the Ctrl+Tab switcher overlay for focus. The host may receive
    ; WM_ACTIVATE while the overlay is open (the overlay is +AlwaysOnTop but the
    ; host redraws underneath during preview); refocusing the embedded content
    ; here triggers a tug-of-war with SwitcherFocusGui and stacks paint messages.
    if State.SwitcherGui
        return
    if !State.HostByHwnd.Has(hwnd "")
        return
    host := State.HostByHwnd[hwnd ""]
    if host.activeTabId != "" && host.tabRecords.Has(host.activeTabId) {
        record := host.tabRecords[host.activeTabId]
        FocusEmbeddedContent(host.hwnd, record.contentHwnd)
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
    prev := DetectHiddenWindows(true)
    try {
        WinGetClientPos(,, &w,, "ahk_id " hwnd)
        DetectHiddenWindows(prev)
        return w
    } catch {
        DetectHiddenWindows(prev)
        return 0
    }
}

; Returns the client-area height of a window.
GetClientHeight(hwnd) {
    prev := DetectHiddenWindows(true)
    try {
        WinGetClientPos(,,, &h, "ahk_id " hwnd)
        DetectHiddenWindows(prev)
        return h
    } catch {
        DetectHiddenWindows(prev)
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
    ; +0x02000000 = WS_CLIPCHILDREN. Without it, the host's WM_ERASEBKGND paints the
    ; host background over the area of any embedded child window. Foreign WPF content
    ; uses DirectComposition: it will not auto-recover from GDI damage, so its last
    ; presented frame gets overwritten by white and the app only repaints on a size
    ; change or user interaction. Clipping children out of the parent's paint pipeline
    ; prevents that.
    guiOpts := "+Resize +MinSize" Config.HostMinWidth "x" Config.HostMinHeight " +0x02000000"
    host.gui := Gui(guiOpts, title)
    host.gui.BackColor := Config.ThemeBackground
    host.gui.MarginX := 0
    host.gui.MarginY := 0
    host.gui.SetFont("s" Config.ThemeFontSize " c" Config.ThemeWindowText, Config.ThemeFontName)
    host.gui.OnEvent("Close", HostGuiClosed.Bind(host))
    host.gui.OnEvent("Size", HostGuiResized.Bind(host))

    tabBarH := Config.HeaderHeight
    tabBarY := (Config.TabPosition = "bottom") ? Config.HostHeight - tabBarH : 0
    host.tabBarBg := host.gui.Add("Text", "x0 y0 w" Config.HostWidth " h" tabBarH " Background" Config.ThemeTabBarBg, "")
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
        if host.HasProp("tabBarHBitmap") && host.tabBarHBitmap {
            DllCall("DeleteObject", "UPtr", host.tabBarHBitmap)
            host.tabBarHBitmap := 0
        }
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
    if host.HasProp("isLayingOut") && host.isLayingOut
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
    ; Skip if already embedded — check by id AND by hwnd (title change can produce a different id)
    if State.MainHost.tabRecords.Has(freshCandidate.id)
        return ""
    for _, rec in State.MainHost.tabRecords {
        if rec.topHwnd = freshCandidate.topHwnd || rec.contentHwnd = freshCandidate.contentHwnd
            return ""
    }
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
            ; Only discard if the window is gone — slow-starting WPF apps can take several
            ; seconds to set a title and should not be silently dropped while still alive.
            if !IsWindowExists(pending.candidate.topHwnd)
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
    ; Only call Show when the window is actually hidden — avoids repositioning an already-visible window
    if Config.ShowOnlyWhenTabs && State.MainHost.tabOrder.Length >= 1 {
        if !(GetWindowLongPtrValue(State.MainHost.hwnd, -16) & 0x10000000) {  ; WS_VISIBLE
            ; Set title before Show so tiling WMs (e.g. GlazeWM) see the correct title on EVENT_OBJECT_SHOW
            UpdateHostTitle(State.MainHost)
            State.MainHost.gui.Show()  ; Show() activates by default; intentional for first appearance
        }
        ; If already visible don't force-activate — user may have focus elsewhere
    }
    LayoutTabButtons(State.MainHost)
    ShowOnlyActiveTab(State.MainHost)
    UpdateHostTitle(State.MainHost)
    RedrawAnyWindow(State.MainHost.hwnd)
    ; Deferred layout + redraw: WM_SIZE from Show() may not have been processed yet when the
    ; calls above run, causing a blank tab bar or wrong embed size on first appearance.
    ; Both fire after the message loop has settled with the correct window dimensions.
    SetTimer(() => (State.MainHost ? ShowOnlyActiveTab(State.MainHost) : 0), -148)
    SetTimer(() => (State.MainHost ? DrawTabBar(State.MainHost) : 0), -150)
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
    if (event != 0x8002 && event != 0x800C && event != 0x8018 && event != 0x0003)
        return
    if (idObject != 0 || !hwnd)
        return
    ; EVENT_SYSTEM_FOREGROUND: dispatched separately so it doesn't fall through
    ; the SHOW/NAMECHANGE/UNCLOAK logic below.
    if (event = 0x0003) {
        OnForegroundChanged(hwnd)
        return
    }
    ; NAMECHANGE on already-tracked tab: update the title label and redraw the tab bar.
    ; Do NOT call SelectTab here — switching the active tab on every title change (e.g. Notepad
    ; adding a '*' prefix when unsaved) would steal focus from whichever tab the user is in.
    if (event = 0x800C) {
        for host in GetAllHosts() {
            for tabId, record in host.tabRecords {
                if (record.topHwnd = hwnd || record.contentHwnd = hwnd) {
                    newTitle := GetPreferredTabTitle(record)
                    if newTitle != "" && newTitle != record.title {
                        record.title := newTitle
                        UpdateRecordTitleCache(record)
                        DrawTabBar(host)
                        if tabId = host.activeTabId
                            UpdateHostTitle(host)
                    }
                    return
                }
            }
        }
    }
    if (event = 0x8002) {
        for host in GetAllHosts() {
            for tabId, record in host.tabRecords {
                if (record.topHwnd = hwnd && record.topHwnd != record.contentHwnd) {
                    SetTimer(ReEmbedTab.Bind(host, tabId), -20)
                    return
                }
            }
        }
    }
    ; When KeepAboveTabApps is on: reparent an owned dialog belonging to a tracked
    ; pid to the host and bump the host forward once. Reparenting leverages the
    ; OS owner rule to keep the dialog above the host structurally, so the
    ; enforcer tick doesn't have to catch the "buried popup" case every frame.
    ; Only on SHOW (0x8002): NAMECHANGE/UNCLOAK must not trigger. Match on PID
    ; (not processName) so multiple instances of the same exe bump the right host.
    ; After the bump, return to avoid OnWindowCreated trying to embed the dialog
    ; as a tab.
    ;
    ; Delay the bump by ~80ms so the dialog's own ShowWindow/activation has
    ; completed before we reorder. EVENT_OBJECT_SHOW fires early in that sequence;
    ; bumping synchronously wins the race for an instant and loses it a frame later.
    if (event = 0x8002) && Config.KeepAboveTabApps {
        try {
            newPid := WinGetPID("ahk_id " hwnd)
            if newPid != "" {
                ownerHwnd := DllCall("GetWindow", "ptr", hwnd, "uint", 4, "ptr")  ; GW_OWNER
                for host in GetAllHosts() {
                    for tabId, record in host.tabRecords {
                        if (record.pid = newPid
                            && record.topHwnd != hwnd
                            && record.contentHwnd != hwnd) {
                            if Config.KeepAboveTabAppsDebug
                                DbgZ("SHOW tracked-pid owned win=" DbgZDescribe(hwnd) " owner=" DbgZDescribe(ownerHwnd))
                            ; Reparent the dialog's owner-chain to the host when it has any
                            ; owner in the same process. WPF apps typically set dialog.Owner
                            ; to Application.MainWindow, not to the embedded content window,
                            ; so a strict match (owner = topHwnd || contentHwnd) never fires
                            ; for those apps. Matching "any owner in same pid" covers both
                            ; cases: dialog owned by the embedded tab AND dialog owned by the
                            ; app's main window. Never reparent if owner is 0 (means it's a
                            ; top-level with no owner - could be the main window itself) or
                            ; if owner is already the host.
                            if ownerHwnd && ownerHwnd != host.hwnd {
                                try {
                                    ownerPid := WinGetPID("ahk_id " ownerHwnd)
                                    if ownerPid = newPid {
                                        DllCall("SetWindowLongPtr", "ptr", hwnd
                                            , "int", -8, "ptr", host.hwnd, "ptr")  ; GWLP_HWNDPARENT
                                    }
                                }
                            }
                            SetTimer(BumpHostToFront.Bind(host), -80)
                            return
                        }
                    }
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
    try {
        if State.SwitcherGui && hwnd = State.SwitcherGui.Hwnd
            return
    } catch {
        State.SwitcherGui := ""  ; Gui destroyed before State was cleared
    }
    ; Never embed our own host windows
    for host in GetAllHosts() {
        if host.hwnd = hwnd
            return
    }
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
                ; Destroy empty popout — CloseTabDeferredUpdate handles this for StackTabs-initiated
                ; closes, but external window closes bypass that path and leave a blank popout visible.
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
                ; Build hwnd index so we can detect already-tracked windows whose id changed (e.g. title change)
                trackedHwnds := Map()
                for _, rec in host.tabRecords {
                    trackedHwnds[rec.topHwnd ""] := true
                    trackedHwnds[rec.contentHwnd ""] := true
                }
                currentHwnds := Map()
                for candidate in candidates {
                    currentIds[candidate.id] := true
                    currentHwnds[candidate.topHwnd ""] := true

                    if host.tabRecords.Has(candidate.id) {
                        if UpdateTrackedTab(host, candidate.id, candidate)
                            structureChanged := true
                        continue
                    }

                    ; Guard: skip if this hwnd is already tracked under a different id (e.g. title changed)
                    if trackedHwnds.Has(candidate.topHwnd "") || trackedHwnds.Has(candidate.contentHwnd "")
                        continue

                    if !State.PendingCandidates.Has(candidate.id) {
                        TryStackOrPending(host, candidate)
                        if host.tabRecords.Has(candidate.id)
                            structureChanged := true
                        continue
                    }

                    ; Already pending: refresh candidate metadata only, do NOT reset firstSeen
                    State.PendingCandidates[candidate.id].candidate := candidate
                }

                ; Remove pending candidates whose window is gone — use hwnd not id,
                ; so a title change (which changes the id) doesn't reset the pending timer.
                stalePending := []
                for tabId, pending in State.PendingCandidates {
                    if !currentHwnds.Has(pending.candidate.topHwnd "")
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

    ; Fallback: destroy any popout hosts that ended up empty (e.g. stale-tab cleanup left them blank).
    emptyPopouts := []
    for h in State.PopoutHosts {
        if h.tabOrder.Length = 0
            emptyPopouts.Push(h)
    }
    for h in emptyPopouts {
        for i, ph in State.PopoutHosts {
            if ph = h {
                State.PopoutHosts.RemoveAt(i)
                break
            }
        }
        if h.hwnd
            State.HostByHwnd.Delete(h.hwnd "")
        InvalidateHostsCache()
        h.gui.Destroy()
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
        ; Never embed our own host windows — prevents StackTabs from stacking itself
        for host in GetAllHosts() {
            if host.hwnd = topHwnd
                return ""
        }
        title := WinGetTitle("ahk_id " topHwnd)
        if !DllCall("IsWindowVisible", "ptr", topHwnd)
            return ""
        if Config.WindowTitleMatches.Length = 0
            return ""  ; No match patterns configured
        ; Match1="" in config is a sentinel meaning "match windows with no title / whitespace-only title".
        ; All other patterns do a case-insensitive substring match against the title.
        matched := false
        for pat in Config.WindowTitleMatches {
            if pat = '""' {
                if Trim(title) = ""
                    matched := true
            } else if InStr(title, pat, false) {
                matched := true
            }
            if matched
                break
        }
        if !matched
            return ""

        processName := WinGetProcessName("ahk_id " topHwnd)
        pid := WinGetPID("ahk_id " topHwnd)
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
            pid: pid,
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
        if pat = '""' {
            if Trim(title) = ""
                titleMatches := true
        } else if (title != "") && InStr(title, pat, false) {
            titleMatches := true
        }
        if titleMatches
            break
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
        pid: candidate.pid,
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
    record.pid := candidate.pid
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
    ; Hide host immediately when closing the last tab — CloseWindowReliably is synchronous
    ; and can block for up to 2000ms per message, so the host would stay visible otherwise.
    if !host.isPopout && Config.ShowOnlyWhenTabs && host.tabOrder.Length = 1
        host.gui.Hide()
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
        try WinActivate("ahk_id " host.hwnd)
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

; Re-hides topHwnd when it re-appears while stacked; shows the host if it was tray-hidden.
; Does NOT restore a minimized host — the user minimized it intentionally; let HostGuiResized
; re-layout on restore (WinRestore here would pop up the host against the user's intent).
ReEmbedTab(host, tabId, *) {
    if !host.tabRecords.Has(tabId)
        return
    record := host.tabRecords[tabId]
    if IsWindowExists(record.topHwnd) && record.topHwnd != record.contentHwnd
        DllCall("ShowWindow", "ptr", record.topHwnd, "int", SW_HIDE)
    ; If host is gone or minimized: topHwnd is now hidden; bail out.
    ; HostGuiResized(minMax=0) will call ShowOnlyActiveTab when the user restores the host.
    if !IsWindowExists(host.hwnd) || DllCall("IsIconic", "ptr", host.hwnd, "int")
        return
    if Config.ShowOnlyWhenTabs && !WinExist("ahk_id " host.hwnd)
        host.gui.Show()
    if IsWindowExists(record.contentHwnd) {
        if DllCall("GetParent", "ptr", record.contentHwnd, "ptr") != host.clientHwnd
            AttachTrackedWindow(host, tabId)
    }
    LayoutTabButtons(host)
    ShowOnlyActiveTab(host)
    RedrawAnyWindow(host.hwnd)
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

    ; Hide content window immediately after reparent — SetParent makes WS_VISIBLE children
    ; briefly visible, which causes a background flash in the host. Hide first, then position.
    DllCall("ShowWindow", "ptr", hwnd, "int", SW_HIDE)

    ; Position at final content rect so when ShowOnlyActiveTab shows it there's no resize glitch.
    GetEmbedRect(host, &areaX, &areaY, &areaW, &areaH)
    areaX += 1
    areaY += 1
    areaW -= 2
    areaH -= 2
    flags := SWP_FRAMECHANGED | SWP_NOZORDER | SWP_NOACTIVATE
    DllCall("SetWindowPos", "ptr", hwnd, "ptr", 0, "int", areaX, "int", areaY, "int", areaW, "int", areaH, "uint", flags)
    try SendMessage(0x000B, 1, 0,, "ahk_id " hwnd,,,, 500)  ; WM_SETREDRAW TRUE
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

        ; Only include SWP_SHOWWINDOW when restoring — avoids briefly flashing the window
        ; on screen (at its original position) right before we hide it on close.
        flags := SWP_FRAMECHANGED | (restoreWindow ? SWP_SHOWWINDOW : 0)
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

; Repositions tab bar controls to match the current window size, then redraws.
; Called on WM_SIZE and after theme changes. Separating positioning from drawing
; avoids moving controls on every redraw (which can cause flicker).
LayoutTabButtons(host, windowWidth := 0, windowHeight := 0) {
    if !host || !host.gui || !host.hwnd || !IsWindowExists(host.hwnd)
        return
    w := windowWidth > 0 ? windowWidth : GetClientWidth(host.hwnd)
    h := windowHeight > 0 ? windowHeight : GetClientHeight(host.hwnd)
    if !w || !h
        return
    tabBarH := Config.HeaderHeight
    tabBarY := (Config.TabPosition = "bottom") ? h - tabBarH : 0
    if host.HasProp("tabBarBg") && host.tabBarBg
        host.tabBarBg.Move(0, tabBarY, w, tabBarH)
    if host.HasProp("tabCanvas") && host.tabCanvas
        host.tabCanvas.Move(0, tabBarY, w, tabBarH)
    DrawTabBar(host)
}

; Sets active tab, shows its content, updates host title.
SelectTab(host, tabId, *) {
    if !host.tabRecords.Has(tabId)
        return

    ; When the switcher preview already landed on this tab, SwitcherActivate -> SelectTab
    ; would otherwise run a third full ShowOnlyActiveTab within ~100ms. Skip the redundant
    ; layout work but still focus the embedded content so typing lands correctly on commit.
    alreadyActive := (host.activeTabId = tabId)
    host.activeTabId := tabId
    if !alreadyActive {
        ShowOnlyActiveTab(host)
        UpdateHostTitle(host)
    }
    if IsWindowExists(host.tabRecords[tabId].contentHwnd)
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
; scheduleDeferred: pass false when called from DeferredRepaintCheck to avoid infinite chain.
ShowOnlyActiveTab(host, scheduleDeferred := true) {
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
        ; If the real parent is not our host (e.g. after minimize/restore or app reparenting),
        ; SetWindowPos uses client coords relative to the wrong window — often top-left of the
        ; original app. Re-attach before positioning.
        if DllCall("GetParent", "ptr", record.contentHwnd, "ptr") != host.clientHwnd
            AttachTrackedWindow(host, tabId)
        ; Position without SWP_NOCOPYBITS - avoids erasing pixels before the window repaints.
        flags := SWP_FRAMECHANGED | SWP_NOZORDER | SWP_NOACTIVATE
        DllCall("SetWindowPos", "ptr", record.contentHwnd, "ptr", 0
            , "int", areaX, "int", areaY, "int", areaW, "int", areaH
            , "uint", flags)
        DllCall("ShowWindow", "ptr", record.contentHwnd, "int", SW_SHOWNOACTIVATE)
        ; Immediately queue a WM_PAINT on the content hwnd + all its children. Without this,
        ; Windows treats SW_SHOW as "restore the previously valid bitmap" and no paint is
        ; generated, so the host's background (= white) shows through until the 50ms deferred
        ; check forces a sync repaint. Async flags only (no UPDATENOW here): a synchronous
        ; cross-process dispatch inside ShowOnlyActiveTab pumps messages and can land a pending
        ; tab canvas WM_PAINT in the default Static WndProc, corrupting the tab bar.
        DllCall("RedrawWindow", "ptr", record.contentHwnd, "ptr", 0, "ptr", 0
            , "uint", 0x0001 | 0x0004 | 0x0080)  ; INVALIDATE|ERASE|ALLCHILDREN
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
        if IsWindowExists(record.contentHwnd) {
            if DllCall("GetParent", "ptr", record.contentHwnd, "ptr") != host.clientHwnd
                AttachTrackedWindow(host, tabId)
            DllCall("ShowWindow", "ptr", record.contentHwnd, "int", SW_HIDE)
        }
    }
    DrawTabBar(host)
    host.lastRefreshActiveTabId := host.activeTabId
    ; Deferred re-layout at 50ms: re-attempts ShowOnlyActiveTab in case content wasn't ready
    ; (e.g. WPF rendering pipeline hasn't settled). Debounced so rapid tab switches don't stack.
    if scheduleDeferred && host.activeTabId != "" {
        if host.HasProp("deferredRepaintFn") && host.deferredRepaintFn
            SetTimer(host.deferredRepaintFn, 0)
        host.deferredRepaintFn := DeferredRepaintCheck.Bind(host)
        SetTimer(host.deferredRepaintFn, -50)
    }
}

DrawTabBar(host) {

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

    GdipCreateOffscreenBitmap(tabBarW, tabBarH, &pBitmap, &pGraphics)
    if !pBitmap || !pGraphics {
        if pGraphics
            DllCall("gdiplus\GdipDeleteGraphics", "UPtr", pGraphics)
        if pBitmap
            DllCall("gdiplus\GdipDisposeImage", "UPtr", pBitmap)
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
        ApplyBitmapToCanvas(host, pBitmap, pGraphics)
        return
    }

    alignHVal := (Config.TabTitleAlignH = "left") ? 0 : (Config.TabTitleAlignH = "right") ? 2 : 1
    alignVVal := (Config.TabTitleAlignV = "top") ? 0 : (Config.TabTitleAlignV = "bottom") ? 2 : 1

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

    ; Create three tab background brushes once, reuse in loop, delete after
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

        ; Tab background (rounded corners)
        GdipFillRoundRectWithBrush(pGraphics, x, tabOffsetY, tabWidth, Config.TabHeight, Config.TabCornerRadius, pTabBgBrush, HexToARGB(Config.ThemeTabBarBg))

        ; Active indicator strip
        if isActive && Config.TabIndicatorHeight > 0 {
            indicColor := HexToARGB(Config.ThemeTabIndicatorColor != ""
                ? Config.ThemeTabIndicatorColor : Config.ThemeTabActiveBg)
            indicY := (Config.TabPosition = "bottom")
                ? tabOffsetY
                : tabOffsetY + Config.TabHeight - Config.TabIndicatorHeight
            indicW := Max(1, tabWidth - 8)
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

        ; Popout / merge icon
        if Config.ShowPopoutButton {
            iconText := host.isPopout ? Config.IconMerge : Config.IconPopout
            GdipDrawStringSimple(pGraphics, iconText,
                x + titleWidth, tabOffsetY, Config.PopoutButtonWidth, Config.TabHeight,
                iconColor, Config.ThemeIconFont, Config.ThemeIconFontSize, false)
        }

        ; Close icon
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

    ; Scroll arrows
    if needScroll && host.tabScrollOffset > 0 {
        GdipDrawStringSimple(pGraphics, Chr(0xE76B),
            Config.HostPadding, tabOffsetY, arrowW, Config.TabHeight,
            HexToARGB(Config.ThemeTabActiveText), Config.ThemeIconFont, Config.ThemeIconFontSize, false)
    }
    if needScroll && host.tabScrollOffset < host.tabScrollMax {
        GdipDrawStringSimple(pGraphics, Chr(0xE76C),
            Config.HostPadding + arrowW + (drawEnd - drawStart + 1) * (tabWidth + Config.TabGap) - Config.TabGap,
            tabOffsetY, arrowW, Config.TabHeight,
            HexToARGB(Config.ThemeTabActiveText), Config.ThemeIconFont, Config.ThemeIconFontSize, false)
    }

    ApplyBitmapToCanvas(host, pBitmap, pGraphics)
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
    UpdateHostIcon(host)
}

; Computes content area rect (x, y, w, h) for embedded windows.
GetEmbedRect(host, &x, &y, &w, &h) {

    padBottom := (Config.HostPaddingBottom >= 0) ? Config.HostPaddingBottom : Config.HostPadding
    x := Config.HostPadding

    if Config.TabPosition = "bottom"
        y := Config.HostPadding
    else
        y := Config.HeaderHeight + Config.HostPadding

    if host.hwnd && IsWindowExists(host.hwnd) {
        try {
            WinGetClientPos(,, &clientW, &clientH, "ahk_id " host.hwnd)
            ; Guard: if WM_SIZE hasn't settled yet (window just shown), fall through to Config defaults
            if clientW > 0 && clientH > 0 {
                w := Max(200, clientW - (Config.HostPadding * 2))
                if Config.TabPosition = "bottom"
                    h := Max(140, clientH - y - Config.HeaderHeight - padBottom)
                else
                    h := Max(140, clientH - y - padBottom)
                return
            }
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
    StopZOrderEnforcer()

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
    ; Release stored tab bar HBITMAPs
    for host in GetAllHosts() {
        if host.HasProp("tabBarHBitmap") && host.tabBarHBitmap {
            DllCall("DeleteObject", "UPtr", host.tabBarHBitmap)
            host.tabBarHBitmap := 0
        }
    }
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
; Falls back to process name for truly nameless windows (e.g. matched via Match1="").
; Uses GetWindowTextW directly so it works even when topHwnd is hidden after embedding.
GetPreferredTabTitle(record) {
    title := GetWindowTitleDirect(record.topHwnd)
    if title != ""
        return title
    title := GetWindowTitleDirect(record.contentHwnd)
    if title != ""
        return title
    return record.HasProp("processName") ? record.processName : ""
}

; Raises the host to the top of the non-topmost z-order without activating it.
; When KeepAboveTabApps is on, dialogs in the tracked pid have already been
; reparented to the host (see OnWinEventProc's SHOW branch), so a single
; HWND_TOP insert is enough: the OS owner rule keeps the dialog above the host
; on every subsequent activation.
; Uses SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE so the window is not moved,
; resized, or focused.
BumpHostToFront(host, reason := "") {
    if !host || !host.hwnd || !IsWindowExists(host.hwnd)
        return
    if Config.KeepAboveTabAppsDebug
        DbgZBefore("BUMP_TOP", host.hwnd, 0, reason)
    ret := DllCall("SetWindowPos", "ptr", host.hwnd,
        "ptr", 0,        ; HWND_TOP
        "int", 0, "int", 0, "int", 0, "int", 0,
        "uint", 0x0013)  ; SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE
    err := A_LastError
    if Config.KeepAboveTabAppsDebug
        DbgZAfter("BUMP_TOP", host.hwnd, 0, ret, err)
}

; Foreground-change handler (driven by EVENT_SYSTEM_FOREGROUND).
; When KeepAboveTabApps is enabled, each time a top-level window belonging to a
; tracked tab's process becomes the foreground window we run the enforcer
; immediately so the fix is applied within a frame of the event rather than
; waiting for the next enforcer tick. Delegates the actual decision (slot
; below popup vs. bump to top) to EnforceHostZOrder, which inspects the full
; z-order instead of guessing from the single foreground hwnd.
;
; Skips StackTabs hosts themselves to prevent self-bump loops.
OnForegroundChanged(hwnd) {
    if !Config.KeepAboveTabApps
        return
    if !hwnd
        return
    for h in GetAllHosts() {
        if h.hwnd = hwnd {
            if Config.KeepAboveTabAppsDebug
                DbgZ("FG self-host (ignored) " DbgZDescribe(hwnd))
            return
        }
    }
    try {
        fgPid := WinGetPID("ahk_id " hwnd)
        if fgPid = ""
            return
        if Config.KeepAboveTabAppsDebug
            DbgZ("FG " DbgZDescribe(hwnd))
        ; Whether the foreground window is owned (popup/dialog) or top-level
        ; (main shell). Used to remember popups and detect close transitions.
        fgOwner := DllCall("GetWindow", "ptr", hwnd, "uint", 4, "ptr")  ; GW_OWNER
        for host in GetAllHosts() {
            if !host.hwnd || !IsWindowExists(host.hwnd)
                continue
            matched := false
            for tabId, record in host.tabRecords {
                if record.pid = fgPid {
                    matched := true
                    break
                }
            }
            if matched {
                if Config.KeepAboveTabAppsDebug
                    DbgZ("  FG matched host=" DbgZHex(host.hwnd) " -> Enforce(reason=fg)")
                EnforceHostZOrder(host, "fg")
                ; Focus-follow: when the app's popup/dialog is dismissed, Windows
                ; activates the popup's managed Owner (typically the app's main
                ; shell) rather than the host. Detect the transition
                ;   tracked popup FG  ->  tracked main shell FG
                ; and, if the popup is gone, send focus to the host that was
                ; hosting the tab the popup belonged to. This mirrors the
                ; "work on host, popup in front, close popup, keep working on
                ; host" mental model that KeepAboveTabApps aims for.
                if fgOwner {
                    ; FG is a tracked popup/dialog; remember for close detection.
                    State.LastActiveTrackedPopup := { host: host, hwnd: hwnd }
                } else {
                    ; FG is a tracked main shell. If the last tracked popup
                    ; for this same host is destroyed or hidden, the user just
                    ; closed it; redirect focus to the host.
                    prev := State.LastActiveTrackedPopup
                    if IsObject(prev) && prev.host = host {
                        popupGone := !IsWindowExists(prev.hwnd)
                            || !DllCall("IsWindowVisible", "ptr", prev.hwnd, "int")
                        if popupGone {
                            if Config.KeepAboveTabAppsDebug
                                DbgZ("  popup " DbgZHex(prev.hwnd) " gone -> activating host " DbgZHex(host.hwnd))
                            ; Small delay lets the app's own post-close
                            ; activation settle before we override it.
                            SetTimer(ActivateHostAfterPopupClose.Bind(host), -50)
                        }
                        State.LastActiveTrackedPopup := ""
                    }
                }
            }
        }
    }
}

; Called from a timer after a tracked popup has closed and foreground has
; landed on the app's main shell. Activates the host so the user's focus
; returns to what they were working on rather than the app's main shell.
ActivateHostAfterPopupClose(host) {
    if !host || !host.hwnd || !IsWindowExists(host.hwnd)
        return
    if !DllCall("IsWindowVisible", "ptr", host.hwnd, "int")
        return
    try WinActivate("ahk_id " host.hwnd)
    if Config.KeepAboveTabAppsDebug
        DbgZ("ActivateHostAfterPopupClose host=" DbgZHex(host.hwnd))
}

; Positions the host directly below `aboveHwnd` in the z-order without moving,
; resizing, or activating either window. Used so a foreground popup stays on
; top while the host slots in right beneath it (keeping the app's main shell
; below the host).
SlotHostBelow(host, aboveHwnd, reason := "") {
    if !host || !host.hwnd || !IsWindowExists(host.hwnd)
        return
    if !IsWindowExists(aboveHwnd)
        return
    if Config.KeepAboveTabAppsDebug
        DbgZBefore("SLOT_BELOW", host.hwnd, aboveHwnd, reason)
    ret := DllCall("SetWindowPos", "ptr", host.hwnd, "ptr", aboveHwnd
        , "int", 0, "int", 0, "int", 0, "int", 0
        , "uint", 0x0013)  ; SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE
    err := A_LastError
    if Config.KeepAboveTabAppsDebug
        DbgZAfter("SLOT_BELOW", host.hwnd, aboveHwnd, ret, err)
}

; Persistent z-order enforcer. Walks the top-level window list every
; EnforcerIntervalMs and nudges the host whenever the invariant is violated:
;
;     host is above every visible unowned window of any tracked tab's pid,
;     and below every visible owned (popup/dialog) window of those pids.
;
; Event-driven correction alone (via EVENT_SYSTEM_FOREGROUND) misses the case
; where an app calls BringWindowToTop / SetWindowPos(HWND_TOP) on its main
; shell without producing a foreground change. That is exactly the pattern
; behind reports of the host dropping to the back on every click after a
; popup closes: the app reorders its own group but never re-foregrounds, so
; no single-shot event hook can cover it. A persistent low-frequency poll
; does. All corrections use SWP_NOACTIVATE - focus is never touched.
;
; Interval: 16ms = roughly one 60Hz display frame. Apps that actively raise
; their own group on every user interaction (BringWindowToTop without a
; foreground change) will otherwise cause their main shell to be visible
; for one to two frames before the next correction. At 16ms the correction
; lands in the same frame as the app's raise, so the user at worst sees a
; single-frame flash. Typical walk is 2-3 windows (stops at the host or the
; first violating window), so CPU cost is well under 1%.
StartZOrderEnforcer() {
    if State.ZOrderEnforcerFn
        return
    State.ZOrderEnforcerFn := EnforceZOrderTick
    SetTimer(State.ZOrderEnforcerFn, 16)
}

StopZOrderEnforcer() {
    if State.ZOrderEnforcerFn {
        SetTimer(State.ZOrderEnforcerFn, 0)
        State.ZOrderEnforcerFn := ""
    }
}

EnforceZOrderTick(*) {
    if !Config.KeepAboveTabApps {
        StopZOrderEnforcer()
        return
    }
    if State.IsCleaningUp
        return
    for host in GetAllHosts()
        EnforceHostZOrder(host)
}

; Walks the full visible top-level z-order in one pass and classifies every
; window belonging to a tracked pid as either main-shell (no owner) or popup
; (has an owner), and whether it is above or below the host. Two violations
; are possible:
;
;   A. A tracked main-shell sits above the host (the app group-raised).
;      Repair: slot the host below the lowest tracked popup we saw above
;      ourselves (keeps that popup on top), or bump to HWND_TOP if there
;      was none.
;
;   B. A tracked popup sits below the host. This happens when the app opens
;      an owned dialog without activating it: the OS places the popup above
;      its owner (the main shell) to satisfy owner-group rules, but because
;      the host is already above the main shell the popup ends up sandwiched
;      below the host - invisible to the user. Repair: slot the host just
;      below that popup so it becomes visible.
;
; Violation B takes priority because it is the one the user cannot work
; around (the popup is literally hidden). If both are present in the same
; frame, fixing B first will expose the popup; a subsequent tick (60ms
; later) will re-check and fix any remaining main-shell-above-host case.
;
; Only the host is repositioned. No other window's parent, owner, or
; z-order is touched. SWP_NOACTIVATE throughout.
EnforceHostZOrder(host, reason := "tick") {
    if !host || !host.hwnd || !IsWindowExists(host.hwnd)
        return
    ; Only visible hosts need enforcement. A hidden host being "behind"
    ; everything is by design.
    if !DllCall("IsWindowVisible", "ptr", host.hwnd, "int")
        return
    if host.tabRecords.Count = 0
        return
    ; Reentrancy guard. DbgZSnapshot() and other debug writes call
    ; WinGetPID/WinGetTitle, which pump messages and can let a queued
    ; SetTimer(EnforceZOrderTick) or foreground-change callback fire
    ; mid-pass. Without this guard, two enforces see the same
    ; "shell above host" state and each issue their own BumpHostToFront,
    ; so SetWindowPos runs twice for a single violation. The second call
    ; is always a no-op but still costs a round trip to win32k. Keep the
    ; first pass, drop the rest until it returns.
    if State.ZOrderEnforceBusy
        return
    State.ZOrderEnforceBusy := true
    try {
        EnforceHostZOrderImpl(host, reason)
    } finally {
        State.ZOrderEnforceBusy := false
    }
}

EnforceHostZOrderImpl(host, reason) {
    trackedPids := Map()
    for tabId, record in host.tabRecords
        trackedPids[record.pid] := true

    seenHost := false
    lastPopupAboveHost   := 0   ; lowest-z-order tracked popup above the host
    topPopupBelowHost    := 0   ; highest-z-order tracked popup below the host
    shellAboveHost       := false
    trackedSeen          := 0   ; total tracked-pid visible windows scanned
    hwnd := DllCall("GetTopWindow", "ptr", 0, "ptr")
    while hwnd {
        if hwnd = host.hwnd {
            seenHost := true
            hwnd := DllCall("GetWindow", "ptr", hwnd, "uint", 2, "ptr")  ; GW_HWNDNEXT
            continue
        }
        if DllCall("IsWindowVisible", "ptr", hwnd, "int") {
            pid := 0
            try pid := WinGetPID("ahk_id " hwnd)
            if pid && trackedPids.Has(pid) {
                trackedSeen++
                owner := DllCall("GetWindow", "ptr", hwnd, "uint", 4, "ptr")  ; GW_OWNER
                if owner {
                    if !seenHost
                        lastPopupAboveHost := hwnd
                    else if !topPopupBelowHost {
                        topPopupBelowHost := hwnd
                        break   ; first tracked popup below host is all we need
                    }
                } else if !seenHost {
                    shellAboveHost := true
                }
            }
        }
        hwnd := DllCall("GetWindow", "ptr", hwnd, "uint", 2, "ptr")  ; GW_HWNDNEXT
    }

    ; Log decision only when we actually act; no-op ticks would flood the log.
    ; Violation B first: a hidden popup is a user-visible regression.
    if topPopupBelowHost {
        if Config.KeepAboveTabAppsDebug {
            DbgZ("ENFORCE[" reason "] host=" DbgZHex(host.hwnd)
                . " action=SLOT_BELOW_BURIED_POPUP"
                . " popupBelow=" DbgZHex(topPopupBelowHost)
                . " shellAbove=" (shellAboveHost ? "1" : "0")
                . " popupAbove=" DbgZHex(lastPopupAboveHost)
                . " trackedSeen=" trackedSeen)
            DbgZ(DbgZDescribe(topPopupBelowHost) " <-- buried popup")
            DbgZ(DbgZSnapshot())
        }
        SlotHostBelow(host, topPopupBelowHost, "buried_popup")
        if Config.KeepAboveTabAppsDebug
            DbgZ("AFTER:" "`r`n" DbgZSnapshot())
        return
    }
    if shellAboveHost {
        if Config.KeepAboveTabAppsDebug {
            DbgZ("ENFORCE[" reason "] host=" DbgZHex(host.hwnd)
                . " action=" (lastPopupAboveHost ? "SLOT_BELOW_POPUP_ABOVE" : "BUMP_TOP")
                . " popupAbove=" DbgZHex(lastPopupAboveHost)
                . " trackedSeen=" trackedSeen)
            DbgZ(DbgZSnapshot())
        }
        if lastPopupAboveHost
            SlotHostBelow(host, lastPopupAboveHost, "shell_above")
        else
            BumpHostToFront(host, "shell_above")
        if Config.KeepAboveTabAppsDebug
            DbgZ("AFTER:" "`r`n" DbgZSnapshot())
        return
    }
    ; Invariant held; for foreground events (not routine ticks) note it.
    if reason != "tick" && Config.KeepAboveTabAppsDebug
        DbgZ("ENFORCE[" reason "] host=" DbgZHex(host.hwnd) " action=NOOP trackedSeen=" trackedSeen)
}

; ============ Z-ORDER DEBUG LOGGING ============
; All helpers below are gated on Config.KeepAboveTabAppsDebug by their callers
; (or by the flag check inside DbgZ itself). The file is rewritten at startup
; so a single session's output is self-contained.
;
; Log format: each record begins with a "HH:mm:ss.mmm tick=N  " prefix; multi-
; line snapshots are written as separate records so they line up by timestamp.

DbgZPath() {
    return A_ScriptDir "\debug-zorder.log"
}

DbgZ(msg) {
    if !Config.HasProp("KeepAboveTabAppsDebug") || !Config.KeepAboveTabAppsDebug
        return
    try {
        stamp := FormatTime(, "HH:mm:ss") "." Format("{:03d}", A_MSec)
        FileAppend(stamp " tick=" A_TickCount "  " msg "`r`n", DbgZPath(), "UTF-8")
    }
}

; Pretty-print an hwnd as hex, tolerating 0.
DbgZHex(hwnd) {
    return hwnd ? Format("0x{:x}", hwnd) : "0x0"
}

; One-line description of a window. Includes class, trimmed title, pid, exstyle,
; owner hwnd, and visibility. Every lookup is wrapped in try so a hwnd that
; disappears between calls doesn't abort the log line.
DbgZDescribe(hwnd) {
    if !hwnd
        return "hwnd=0x0"
    cls := "" , title := "" , pid := 0 , ex := 0 , ownr := 0 , vis := 0
    try cls := WinGetClass("ahk_id " hwnd)
    try title := WinGetTitle("ahk_id " hwnd)
    try pid := WinGetPID("ahk_id " hwnd)
    try ex := GetWindowLongPtrValue(hwnd, -20)  ; GWL_EXSTYLE
    try ownr := DllCall("GetWindow", "ptr", hwnd, "uint", 4, "ptr")  ; GW_OWNER
    try vis := DllCall("IsWindowVisible", "ptr", hwnd, "int")
    if StrLen(title) > 48
        title := SubStr(title, 1, 45) "..."
    ; Strip newlines from title so the log stays one-record-per-line.
    title := StrReplace(StrReplace(title, "`r", " "), "`n", " ")
    return Format("{} pid={} cls={} title=`"{}`" ex=0x{:x} owner={} vis={}"
        , DbgZHex(hwnd), pid, cls, title, ex, DbgZHex(ownr), vis)
}

; Returns a multi-line snapshot string of up to `limit` visible top-level
; windows, top-to-bottom in z-order. Every entry is a self-contained record.
DbgZSnapshot(limit := 20) {
    out := ""
    count := 0
    hwnd := DllCall("GetTopWindow", "ptr", 0, "ptr")
    while hwnd && count < limit {
        if DllCall("IsWindowVisible", "ptr", hwnd, "int") {
            out .= "  [" count "] " DbgZDescribe(hwnd) "`r`n"
            count++
        }
        hwnd := DllCall("GetWindow", "ptr", hwnd, "uint", 2, "ptr")
    }
    ; Trim trailing CRLF so DbgZ's own CRLF doesn't produce a blank line.
    return RTrim(out, "`r`n")
}

; Log the state just before a SetWindowPos against the host.
DbgZBefore(action, hostHwnd, anchorHwnd, reason) {
    DbgZ(action " BEFORE host=" DbgZHex(hostHwnd)
        . (anchorHwnd ? " anchor=" DbgZHex(anchorHwnd) : "")
        . (reason ? " reason=" reason : ""))
}

; Log the return value + last error just after a SetWindowPos against the host.
DbgZAfter(action, hostHwnd, anchorHwnd, ret, err) {
    DbgZ(action " AFTER  host=" DbgZHex(hostHwnd)
        . (anchorHwnd ? " anchor=" DbgZHex(anchorHwnd) : "")
        . " ret=" ret " lastError=" err)
}

; Initializes the debug log at startup. The header is written *unconditionally*
; so we can always confirm three things without needing the debug flag on:
;   1. the script is running the current build,
;   2. the config was parsed and what values were picked up,
;   3. the script directory is writable.
; The verbose trace (FG/SHOW/ENFORCE/etc.) is still gated on
; Config.KeepAboveTabAppsDebug via DbgZ's own check.
InitZOrderDebugLog() {
    path := DbgZPath()
    try FileDelete(path)
    try {
        hdr := "=== StackTabs z-order debug log ===`r`n"
        hdr .= "started " FormatTime(, "yyyy-MM-dd HH:mm:ss") "`r`n"
        hdr .= "script " A_ScriptFullPath "`r`n"
        hdr .= "pid " ProcessExist() "`r`n"
        hdr .= "KeepAboveTabApps=" (Config.KeepAboveTabApps ? "1" : "0") "`r`n"
        hdr .= "KeepAboveTabAppsDebug=" (Config.KeepAboveTabAppsDebug ? "1" : "0") "`r`n"
        hdr .= "DebugDiscovery=" (Config.DebugDiscovery ? "1" : "0") "`r`n"
        hdr .= "--- raw IniRead diagnostics (from " Config.ConfigPath ") ---`r`n"
        for key in ["KeepAboveTabApps","KeepAboveTabAppsDebug","DebugDiscovery"] {
            raw := ""
            try raw := IniRead(Config.ConfigPath, "General", key, "<missing>")
            hdr .= "  raw[" key "] = `"" raw "`" (len=" StrLen(raw) ")`r`n"
        }
        hdr .= "--- end raw ---`r`n"
        hdr .= "TargetExe=" Config.TargetExe "`r`n"
        hdr .= "Matches="
        for i, m in Config.WindowTitleMatches
            hdr .= (i > 1 ? "|" : "") m
        hdr .= "`r`n"
        hdr .= (Config.KeepAboveTabAppsDebug
            ? "verbose trace: ON (events will follow below)"
            : "verbose trace: OFF - set KeepAboveTabAppsDebug=1 in [General] and restart to enable")
        hdr .= "`r`n===`r`n"
        FileAppend(hdr, path, "UTF-8")
    } catch as err {
        ; If writing the header fails (e.g. read-only install dir), surface it
        ; via the tray tooltip so the user isn't left wondering why the log
        ; never appears. Non-fatal.
        try A_IconTip := "StackTabs: cannot write debug log to " path
    }
    if Config.KeepAboveTabAppsDebug {
        DbgZ("BOOT initial z-order snapshot:")
        DbgZ(DbgZSnapshot())
    }
}

; Forces window to redraw (invalidates and updates).
RedrawAnyWindow(hwnd) {
    if !hwnd || !IsWindowExists(hwnd)
        return

    ; RDW_UPDATENOW is intentionally omitted: it sends WM_PAINT directly to child WndProcs,
    ; bypassing AHK's OnMessage hook. Without it, WM_PAINT is dispatched via the message loop
    ; where OnTabCanvasPaint fires correctly.
    flags := 0x0001 | 0x0004 | 0x0080 | 0x0400  ; INVALIDATE|ERASE|ALLCHILDREN|FRAME
    DllCall("RedrawWindow", "ptr", hwnd, "ptr", 0, "ptr", 0, "uint", flags)
}

; Timer callback: re-attempts layout and redraws 50ms after ShowOnlyActiveTab.
; Fixes blank content areas where the embedded window wasn't ready on the initial call
; (common with WPF whose rendering pipeline needs time to settle after reparenting).
; Passes scheduleDeferred=false to avoid scheduling another deferred check and looping.
DeferredRepaintCheck(host, *) {
    if host.HasProp("deferredRepaintFn")
        host.deferredRepaintFn := ""
    if !host || !host.hwnd || !IsWindowExists(host.hwnd)
        return
    ShowOnlyActiveTab(host, false)
    RedrawAnyWindow(host.hwnd)
    ; Force-repaint the active embedded window via a secondary timer (20 ms from now).
    ; We cannot call RDW_UPDATENOW inline here: it dispatches WM_PAINT cross-process via a
    ; synchronous call, which causes AHK to pump its own message queue while waiting. If a
    ; WM_PAINT for the tab canvas is pending at that moment (just queued by DrawTabBar above),
    ; it would be handled by the default Static WndProc instead of OnTabCanvasPaint, corrupting
    ; the tab bar. Scheduling 20 ms out lets the message loop drain the tab-canvas WM_PAINT
    ; through OnTabCanvasPaint first, so UPDATENOW on the foreign window is safe.
    ;
    ; The timer is tracked on host so rapid tab switches cancel the previous pending UPDATENOW.
    ; Without cancellation, a stale UPDATENOW could fire against the hwnd of an already-replaced
    ; active tab, pumping messages mid-switch and causing blank content.
    if host.activeTabId && host.tabRecords.Has(host.activeTabId) {
        activeRecord := host.tabRecords[host.activeTabId]
        if IsWindowExists(activeRecord.contentHwnd) {
            chw := activeRecord.contentHwnd
            if host.HasProp("pendingUpdateNowFn") && host.pendingUpdateNowFn
                SetTimer(host.pendingUpdateNowFn, 0)
            host.pendingUpdateNowFn := _StillActiveUpdateNow.Bind(host, chw)
            SetTimer(host.pendingUpdateNowFn, -20)
        }
    }
}

; Secondary UPDATENOW callback scheduled by DeferredRepaintCheck.
; Validates chw is still the active tab's content before firing the cross-process WM_PAINT,
; so a rapid switch that replaces the active tab skips a stale repaint cleanly.
_StillActiveUpdateNow(host, chw, *) {
    if host.HasProp("pendingUpdateNowFn")
        host.pendingUpdateNowFn := ""
    if !host || !host.hwnd || !IsWindowExists(host.hwnd)
        return
    if !host.activeTabId || !host.tabRecords.Has(host.activeTabId)
        return
    if host.tabRecords[host.activeTabId].contentHwnd != chw
        return
    if !IsWindowExists(chw)
        return
    ; Flag correction: 0x0400 is RDW_FRAME, not RDW_ALLCHILDREN (0x0080). The original
    ; comment misnamed the value, which meant UPDATENOW never descended into child hwnds.
    ; WPF windows host their render surface in an HwndSource child, so without ALLCHILDREN
    ; the visual tree never received WM_PAINT and stayed blank. Keep FRAME too so a frame-
    ; changed state still repaints correctly.
    DllCall("RedrawWindow", "ptr", chw, "ptr", 0, "ptr", 0
        , "uint", 0x0001 | 0x0004 | 0x0080 | 0x0100 | 0x0400)  ; INVALIDATE|ERASE|ALLCHILDREN|UPDATENOW|FRAME
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
; Note: respects DetectHiddenWindows — use GetWindowTitleDirect for embedded (hidden) windows.
SafeWinGetTitle(hwnd) {
    try return WinGetTitle("ahk_id " hwnd)
    catch
        return ""
}

; Reads window title via GetWindowTextW directly — works on hidden windows unlike WinGetTitle.
; Use this for already-embedded tabs whose topHwnd has been hidden by StackTabs.
GetWindowTitleDirect(hwnd) {
    if !hwnd
        return ""
    buf := Buffer(1024, 0)
    len := DllCall("GetWindowTextW", "ptr", hwnd, "ptr", buf, "int", 512, "int")
    return len > 0 ? StrGet(buf) : ""
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
        tabBarY := 0
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
        tabBarY := 0
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
        tabBarY := 0
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
        tabBarY := 0
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
        tabBarY := 0
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

OnMessage(0x0006, OnWmActivate)   ; WM_ACTIVATE: re-focus embedded content when host regains focus
OnMessage(0x000F, OnTabCanvasPaint)   ; WM_PAINT: BitBlt stored bitmap to canvas DC on every repaint
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

; Stores the rendered HBITMAP on the host for OnTabCanvasPaint to BitBlt on every WM_PAINT.
; Converts the GDI+ bitmap to an HBITMAP and stores it on the host for OnTabCanvasPaint.
; More reliable than STM_SETIMAGE: the stored bitmap persists across any future WM_PAINT
; (e.g. after another window uncovers the tab bar), so tabs always show correct content.
ApplyBitmapToCanvas(host, pBitmap, pGraphics) {
    hBmp := GdipBitmapToHBITMAP(pBitmap)
    GdipCleanupBitmap(pBitmap, pGraphics)
    if !hBmp
        return
    ; Replace the stored bitmap (we own it; OnTabCanvasPaint only reads it)
    if host.HasProp("tabBarHBitmap") && host.tabBarHBitmap
        DllCall("DeleteObject", "UPtr", host.tabBarHBitmap)
    host.tabBarHBitmap := hBmp
    ; Mark the canvas dirty so WM_PAINT fires via the message loop.
    ; Do NOT call UpdateWindow here — it bypasses AHK's OnMessage hook and calls the
    ; STATIC control's default WndProc directly, which has no bitmap set and paints blank.
    if host.HasProp("tabCanvas") && host.tabCanvas
        DllCall("InvalidateRect", "ptr", host.tabCanvas.Hwnd, "ptr", 0, "int", false)
}

; WM_PAINT handler for tab canvas controls.
; BitBlts the stored HBITMAP directly to the canvas DC so the tab bar always shows
; the correct content — even after being uncovered by another window.
OnTabCanvasPaint(wParam, lParam, msg, hwnd) {
    for host in GetAllHosts() {
        if !host.HasProp("tabCanvas") || host.tabCanvas.Hwnd != hwnd
            continue
        ps := Buffer(64, 0)
        hdc := DllCall("BeginPaint", "ptr", hwnd, "ptr", ps, "ptr")
        if hdc {
            rc := Buffer(16, 0)
            DllCall("GetClientRect", "ptr", hwnd, "ptr", rc)
            bmpW := NumGet(rc, 8, "int")
            bmpH := NumGet(rc, 12, "int")
            if host.HasProp("tabBarHBitmap") && host.tabBarHBitmap {
                memDC := DllCall("CreateCompatibleDC", "ptr", hdc, "ptr")
                old := DllCall("SelectObject", "ptr", memDC, "ptr", host.tabBarHBitmap, "ptr")
                DllCall("BitBlt", "ptr", hdc, "int", 0, "int", 0, "int", bmpW, "int", bmpH,
                    "ptr", memDC, "int", 0, "int", 0, "uint", 0x00CC0020)  ; SRCCOPY
                DllCall("SelectObject", "ptr", memDC, "ptr", old)
                DllCall("DeleteDC", "ptr", memDC)
            } else {
                ; Bitmap not ready yet — fill with the tab bar background color so there's no
                ; flicker or garbage pixels visible before the first DrawTabBar completes.
                hex := Config.ThemeTabBarBg  ; "RRGGBB"
                clr := Integer("0x" SubStr(hex,1,2)) | (Integer("0x" SubStr(hex,3,2)) << 8) | (Integer("0x" SubStr(hex,5,2)) << 16)  ; COLORREF = R|G<<8|B<<16
                hbr := DllCall("CreateSolidBrush", "uint", clr, "ptr")
                DllCall("FillRect", "ptr", hdc, "ptr", rc, "ptr", hbr)
                DllCall("DeleteObject", "ptr", hbr)
            }
        }
        DllCall("EndPaint", "ptr", hwnd, "ptr", ps)
        return 0  ; handled — suppress default Static WM_PAINT
    }
    return ""  ; not our canvas — let Windows handle it
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
        tabBarY := 0
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

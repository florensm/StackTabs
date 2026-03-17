; StackTabs - owner-aware embedded window host
; AutoHotkey v2

#Requires AutoHotkey v2.0

; ============ CONFIGURATION ============
; Paths (relative to script directory)
g_ConfigPath     := A_ScriptDir "\config.ini"
g_ConfigExample  := A_ScriptDir "\config.ini.example"
g_SessionPath    := A_ScriptDir "\session.ini"
g_ThemesDir      := A_ScriptDir "\themes"
g_DebugLogPath   := A_ScriptDir "\discovery.txt"
g_WinEventLogPath := A_ScriptDir "\winevents.txt"  ; always-on timing log for window detection
g_DebugDiscovery := false  ; when true, AppendDebugLog writes to discovery.txt on new candidates

; Window title patterns: Match1/Match2/... in config.ini. Window must contain at least one.
g_WindowTitleMatches := []

; Optional EXE filter. Leave blank to match any process.
g_TargetExe := ""

; Shell Hook + Slow Sweep: event-driven discovery; fallback scan interval.
g_SlowSweepInterval := 3000
g_StackDelayMs := 100       ; Minimum wait before stacking; title-stability check provides additional protection
g_StackSwitchDelayMs := 150  ; Delay before switching to newly stacked tab; lets content load to reduce glitch
g_WatchdogMaxMs := 1500
g_TabDisappearGraceMs := 300

; Host window defaults.
g_HostTitle := "StackTabs"
g_HostWidth := 1200
g_HostHeight := 800
g_HostX := -1   ; -1 = no saved position, let Windows decide
g_HostY := -1
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
; Maximum characters shown in a tab label.
g_TabTitleMaxLen      := 60

; Loads config.ini and session.ini into globals; migrates from StackTabs.ini if needed.
LoadConfigFromIni() {
    ; Migrate from StackTabs.ini if config.ini doesn't exist
    if !FileExist(g_ConfigPath) && FileExist(A_ScriptDir "\StackTabs.ini")
        FileCopy(A_ScriptDir "\StackTabs.ini", g_ConfigPath)
    if !FileExist(g_ConfigPath)
        return
    iniPath := g_ConfigPath
    global g_WindowTitleMatches, g_TargetExe, g_SlowSweepInterval, g_StackDelayMs, g_StackSwitchDelayMs, g_WatchdogMaxMs, g_TabDisappearGraceMs, g_DebugDiscovery
    global g_HostTitle, g_HostWidth, g_HostHeight, g_HostMinWidth, g_HostMinHeight, g_HostX, g_HostY
    global g_HostPadding, g_HostPaddingBottom, g_HeaderHeight, g_TabGap, g_MinTabWidth, g_MaxTabWidth, g_TabHeight
    global g_TabSlotMax, g_CloseButtonWidth, g_PopoutButtonWidth, g_TabBarAlignment, g_TabBarOffsetY, g_TabPosition
    global g_ShowOnlyWhenTabs, g_TabIndicatorHeight, g_ActiveTabStyle
    global g_TitleStripPatterns, g_TabTitleMaxLen
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
        g_TabTitleMaxLen := Integer(IniRead(iniPath, "Layout", "TabTitleMaxLen", g_TabTitleMaxLen))
        g_TabPosition := IniRead(iniPath, "Layout", "TabPosition", "top")
        g_TabIndicatorHeight := Integer(IniRead(iniPath, "Layout", "TabIndicatorHeight", "3"))
        g_ActiveTabStyle := Trim(IniRead(iniPath, "Layout", "ActiveTabStyle", "full"))
        ; ShowOnlyWhenTabs: show host only when 1+ tabs; hide to tray when 0 (default). Fallback for old config keys.
        rawVal := IniRead(iniPath, "Layout", "ShowOnlyWhenTabs", IniRead(iniPath, "Layout", "KeepHostAlive", IniRead(iniPath, "Layout", "HideHostWhenEmpty", "1")))
        ; Strip inline comment (; ...) and trim so "1   ; comment" parses as 1
        g_ShowOnlyWhenTabs := (Trim(StrSplit(rawVal, ";")[1]) = "1")
        g_UseCustomTitleBar := (IniRead(iniPath, "Layout", "UseCustomTitleBar", "0") = "1")
        g_TitleBarHeight := Integer(IniRead(iniPath, "Layout", "TitleBarHeight", "28"))
        g_ActiveThemeFile := Trim(IniRead(iniPath, "Theme", "ThemeFile", "dark.ini"))
    }
    if FileExist(g_SessionPath) {
        g_HostX := Integer(IniRead(g_SessionPath, "Session", "WindowX", "-1"))
        g_HostY := Integer(IniRead(g_SessionPath, "Session", "WindowY", "-1"))
        g_HostWidth  := Integer(IniRead(g_SessionPath, "Session", "WindowW", g_HostWidth))
        g_HostHeight := Integer(IniRead(g_SessionPath, "Session", "WindowH", g_HostHeight))
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
    global g_CloseButtonWidth, g_PopoutButtonWidth, g_TabBarAlignment, g_TabBarOffsetY, g_TabPosition, g_TabIndicatorHeight
    global g_ActiveTabStyle
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
    ; TabPosition: when theme doesn't specify, read from config (not g_TabPosition) so we don't carry over
    ; the previous theme's value when switching themes
    rawPos := IniRead(themePath, "Layout", "TabPosition", "")
    g_TabPosition := (rawPos != "") ? Trim(rawPos) : IniRead(g_ConfigPath, "Layout", "TabPosition", "top")
    rawStyle := IniRead(themePath, "Layout", "ActiveTabStyle", "")
    g_ActiveTabStyle := (rawStyle != "") ? Trim(rawStyle) : "full"
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
    if !host || !host.gui || !DllCall("IsWindow", "ptr", host.hwnd)
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
    for _, indic in host.tabSlotIndicators {
        color := (g_ThemeTabIndicatorColor != "") ? g_ThemeTabIndicatorColor : g_ThemeTabActiveBg
        indic.Opt("Background0x" color)
    }
    ; Update font of slot controls (popout/close) â€” they keep old font until we refresh
    for _, popoutBtn in host.tabSlotPopoutButtons {
        popoutBtn.SetFont("s" g_ThemeIconFontSize, g_ThemeIconFont)
        popoutBtn.Opt("c" g_ThemeIconColor)
    }
    for _, closeBtn in host.tabSlotCloseButtons {
        closeBtn.SetFont("s" g_ThemeIconFontSize, g_ThemeIconFont)
        closeBtn.Opt("c" g_ThemeIconColor)
    }
    LayoutTabButtons(host)
    ShowOnlyActiveTab(host)
    UpdateTabButtonStyles(host)
    UpdateHostTitle(host)
    RedrawHostWindow(host)
}

; ============ STATE ============
g_MainHost := ""             ; HostInstance for main window
g_PopoutHosts := []          ; array of HostInstance for popped-out windows
g_PendingCandidates := Map() ; tabId -> {firstSeen, candidate} (main host only)
g_WatchdogTimerActive := false
g_IsCleaningUp := false
g_WinEventDbgCount := 0

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
g_WinEventHooks.Push(DllCall("SetWinEventHook",
    "UInt", 0x8002, "UInt", 0x8018,
    "Ptr",  0,      "Ptr",  g_WinEventHookCallback,
    "UInt", 0,      "UInt", 0, "UInt", 0, "Ptr"))

OnExit(CleanupAll)

; Always write a fresh timing log so window detection can be diagnosed without DebugView.
try FileDelete(g_WinEventLogPath)
WinEventLog("[StackTabs] Started — DebugDiscovery=" (g_DebugDiscovery ? "ON" : "OFF")
    . " WinEventHooks=" g_WinEventHooks.Length
    . " (hook0=" g_WinEventHooks[1] ")")

RefreshWindows()
SetTimer(RefreshWindows, g_SlowSweepInterval)
SetTimer(CheckTabHoverAll, 50)

; ============ HOTKEYS ============
; Win+Shift+T: toggle host visibility (hide when active, show when hidden).
#+t:: {
    global g_MainHost, g_ShowOnlyWhenTabs

    if !g_MainHost
        return

    if WinExist("ahk_id " g_MainHost.hwnd) && WinActive("ahk_id " g_MainHost.hwnd)
        g_MainHost.gui.Hide()
    else {
        ; When ShowOnlyWhenTabs: only show if we have tabs
        if g_ShowOnlyWhenTabs {
            liveCount := 0
            for tabId in g_MainHost.tabOrder {
                if g_MainHost.tabRecords.Has(tabId) && IsWindowExists(g_MainHost.tabRecords[tabId].contentHwnd)
                    liveCount++
            }
            if liveCount = 0
                return
        }
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
        for tabId in host.tabOrder.Clone()
            RemoveTrackedTab(host, tabId, true)
        for i, h in g_PopoutHosts {
            if h = host {
                g_PopoutHosts.RemoveAt(i)
                break
            }
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

; Sends WM_SYSCOMMAND SC_CLOSE and WM_CLOSE to reliably close a window (works with WPF).
CloseWindowReliably(topHwnd, contentHwnd := "") {
    if !contentHwnd
        contentHwnd := topHwnd
    prevHidden := A_DetectHiddenWindows
    DetectHiddenWindows(true)
    try {
        for hwnd in [topHwnd, contentHwnd] {
            if !hwnd || !WinExist("ahk_id " hwnd)
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

; Returns array of main host plus all popout hosts.
GetAllHosts() {
    global g_MainHost, g_PopoutHosts
    hosts := []
    if g_MainHost
        hosts.Push(g_MainHost)
    for h in g_PopoutHosts
        hosts.Push(h)
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
    host.tabButtons := Map()
    host.tabCloseButtons := Map()
    host.tabSlotButtons := []
    host.tabSlotCloseButtons := []
    host.tabSlotPopoutButtons := []
    host.tabSlotIndicators := []
    host.tabIndicators := Map()

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
    host.contentBorderTop := host.gui.Add("Text", "Hidden x0 y0 w0 h1 Background" g_ThemeContentBorder, "")
    host.contentBorderBottom := host.gui.Add("Text", "Hidden x0 y0 w0 h1 Background" g_ThemeContentBorder, "")
    host.contentBorderLeft := host.gui.Add("Text", "Hidden x0 y0 w1 h0 Background" g_ThemeContentBorder, "")
    host.contentBorderRight := host.gui.Add("Text", "Hidden x0 y0 w1 h0 Background" g_ThemeContentBorder, "")
    host.hwnd := host.gui.Hwnd
    host.clientHwnd := host.hwnd
    global g_HostX, g_HostY, g_ShowOnlyWhenTabs
    showOpts := "w" g_HostWidth " h" g_HostHeight
    if !isPopout && g_HostX >= 0 && g_HostY >= 0
        showOpts := "x" g_HostX " y" g_HostY " " showOpts
    if isPopout
        showOpts .= " Hide"  ; Keep hidden until ArrangeHostsSideBySide positions it
    host.gui.Show(showOpts)
    ; Keep hidden when ShowOnlyWhenTabs (host stays in tray until 1+ tabs)
    if !isPopout && g_ShowOnlyWhenTabs
        host.gui.Hide()

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
        if host.HasProp("iconHandle") && host.iconHandle
            DllCall("DestroyIcon", "ptr", host.iconHandle)
        host.gui.Destroy()
    } else {
        host.gui.Hide()
        return true  ; Prevent default close (keep script running in tray)
    }
}

; On host resize: re-layout tabs, update content area, save session (main host only).
HostGuiResized(host, guiObj, minMax, width, height) {
    global g_MainHost, g_SessionPath, g_HostX, g_HostY, g_HostWidth, g_HostHeight
    if minMax = -1
        return
    LayoutTabButtons(host, width, height)
    ShowOnlyActiveTab(host)
    ; Keep globals and session in sync with actual window position/size (main host only)
    if host = g_MainHost && host.hwnd && WinExist("ahk_id " host.hwnd) {
        try {
            WinGetPos(&wx, &wy, &ww, &wh, "ahk_id " host.hwnd)
            g_HostX := wx
            g_HostY := wy
            g_HostWidth := ww
            g_HostHeight := wh
            SetTimer(SaveSessionDeferred, -300)  ; Debounce: save 300ms after last resize
        }
    }
}

; Saves main host position/size to session.ini (debounced after resize).
SaveSessionDeferred(*) {
    global g_MainHost, g_SessionPath
    if !IsObject(g_MainHost) || !g_MainHost.hwnd || !WinExist("ahk_id " g_MainHost.hwnd)
        return
    try {
        WinGetPos(&wx, &wy, &ww, &wh, "ahk_id " g_MainHost.hwnd)
        IniWrite(wx, g_SessionPath, "Session", "WindowX")
        IniWrite(wy, g_SessionPath, "Session", "WindowY")
        IniWrite(ww, g_SessionPath, "Session", "WindowW")
        IniWrite(wh, g_SessionPath, "Session", "WindowH")
    }
}

; Returns true if window has a non-empty title and is not hung.
IsReadyToStack(hwnd) {
    if !hwnd || !DllCall("IsWindow", "ptr", hwnd)
        return false
    title := SafeWinGetTitle(hwnd)
    if (title = "")
        return false
    hung := DllCall("User32.dll\IsHungAppWindow", "Ptr", hwnd, "Int")
    return !hung
}

; Adds candidate to pending map and starts watchdog timer; stacks after delay when title is stable.
TryStackOrPending(host, candidate) {
    global g_MainHost, g_PendingCandidates, g_StackDelayMs, g_WatchdogMaxMs, g_ShowOnlyWhenTabs, g_WatchdogTimerActive
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
        SetTimer(WatchdogCheck, 20)
    }
}

; Timer callback: stacks candidates that passed delay + title-stability; removes stale pending.
WatchdogCheck(*) {
    global g_MainHost, g_PendingCandidates, g_StackDelayMs, g_StackSwitchDelayMs, g_WatchdogMaxMs, g_WatchdogTimerActive, g_ShowOnlyWhenTabs, g_DebugDiscovery
    if !g_MainHost || g_PendingCandidates.Count = 0 {
        g_WatchdogTimerActive := false
        SetTimer(WatchdogCheck, 0)
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
        ; Re-build candidate from scratch so we use fresh, stable metadata (not the snapshot from creation time)
        freshCandidate := BuildCandidateFromTopWindow(item.candidate.topHwnd)
        if !IsObject(freshCandidate)
            continue
        ; Skip if already embedded (could have been stacked by slow sweep between checks)
        if g_MainHost.tabRecords.Has(freshCandidate.id)
            continue
        if CreateTrackedTab(g_MainHost, freshCandidate) {
            lastStackedTabId := freshCandidate.id
            anyStacked := true
            WinEventLog("[Watchdog] T+" A_TickCount " STACKED hwnd=" freshCandidate.topHwnd " elapsed=" (A_TickCount - item.firstSeen) "ms title='" freshCandidate.title "'")
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
    RedrawHostWindow(g_MainHost)
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
    if host.hwnd && WinExist("ahk_id " host.hwnd)
        WinActivate("ahk_id " host.hwnd)
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

; WinEvent hook callback. Range 0x8002-0x8018 is registered; we only act on three events:
;   0x8002 EVENT_OBJECT_SHOW       — window becomes visible (fires before taskbar registration)
;   0x800C EVENT_OBJECT_NAMECHANGE — title set/changed (WPF often sets title after show)
;   0x8018 EVENT_OBJECT_UNCLOAKED  — UWP/WinUI window uncloaks
; idObject=0 (OBJID_WINDOW) means the event is for the window object itself, not a child control.
WinEventProc(hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime) {
    global g_DebugDiscovery, g_WinEventDbgCount
    if (event != 0x8002 && event != 0x800C && event != 0x8018)
        return
    if (idObject != 0 || !hwnd)
        return
    ; Log unconditionally for the first 20 matching events so we can confirm the hook fires
    ; even before DebugDiscovery is enabled. After that, only log when DebugDiscovery is on.
    g_WinEventDbgCount += 1
    if g_DebugDiscovery || g_WinEventDbgCount <= 20 {
        eventName := (event = 0x8002) ? "SHOW" : (event = 0x800C) ? "NAMECHANGE" : "UNCLOAKED"
        title := ""
        try title := WinGetTitle("ahk_id " hwnd)
        WinEventLog("[WinEvent] T+" A_TickCount " #" g_WinEventDbgCount " event=" eventName " hwnd=" hwnd " title='" title "'")
    }
    OnWindowCreated(hwnd)
}

; Shell hook: when a new window is created, try to add it as a tab if it matches.
OnWindowCreated(hwnd) {
    global g_MainHost, g_PendingCandidates, g_IsCleaningUp, g_DebugDiscovery
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
    if !IsObject(candidate) {
        if g_DebugDiscovery {
            title := ""
            try title := WinGetTitle("ahk_id " hwnd)
            WinEventLog("[OnWindowCreated] T+" A_TickCount " hwnd=" hwnd " SKIP(no match) title='" title "'")
        }
        return
    }
    ; Skip if already embedded or pending
    for host in GetAllHosts() {
        if host.tabRecords.Has(candidate.id) {
            if g_DebugDiscovery
                WinEventLog("[OnWindowCreated] T+" A_TickCount " hwnd=" hwnd " SKIP(already tracked)")
            return
        }
    }
    if g_PendingCandidates.Has(candidate.id) {
        if g_DebugDiscovery
            WinEventLog("[OnWindowCreated] T+" A_TickCount " hwnd=" hwnd " SKIP(already pending)")
        return
    }
    ; Skip if content/top already embedded
    for host in GetAllHosts() {
        for tabId, record in host.tabRecords {
            if (record.contentHwnd = candidate.contentHwnd || record.topHwnd = candidate.topHwnd) {
                if g_DebugDiscovery
                    WinEventLog("[OnWindowCreated] T+" A_TickCount " hwnd=" hwnd " SKIP(hwnd in host)")
                return
            }
        }
    }
    WinEventLog("[OnWindowCreated] T+" A_TickCount " hwnd=" hwnd " ACCEPTED title='" candidate.title "'")
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
                if DllCall("IsWindow", "ptr", hwnd)
                    return
                RemoveTrackedTab(host, tabId, false)
                ; Update layout and visibility
                LayoutTabButtons(host)
                ShowOnlyActiveTab(host)
                UpdateHostTitle(host)
                RedrawHostWindow(host)
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
        if !DllCall("IsWindow", "ptr", host.hwnd)
            continue

        structureChanged := false
        currentIds := Map()

        ; Keep existing embedded tabs alive
        for tabId in host.tabOrder {
            if !host.tabRecords.Has(tabId)
                continue
            record := host.tabRecords[tabId]
            if IsWindowExists(record.contentHwnd) {
                record.lastSeenTick := now
                title := GetPreferredTabTitle(record)
                if title != ""
                    record.title := title
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

        if structureChanged {
            LayoutTabButtons(host)
            RedrawHostWindow(host)
        }

        ; Only refresh content/tabs when something changed to avoid flickering
        needsContentRefresh := structureChanged || (host.activeTabId != (host.HasProp("lastRefreshActiveTabId") ? host.lastRefreshActiveTabId : ""))
        if needsContentRefresh {
            host.lastRefreshActiveTabId := host.activeTabId
            ShowOnlyActiveTab(host)
        }
        UpdateHostTitle(host)
    }

    ; When ShowOnlyWhenTabs: hide when 0 tabs, show when 1+ tabs (only if hidden).
    if g_ShowOnlyWhenTabs && g_MainHost && DllCall("IsWindow", "ptr", g_MainHost.hwnd) {
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
    global g_WindowTitleMatches, g_TargetExe

    try {
        if !WinExist("ahk_id " topHwnd)
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
            hierarchySummary: DescribeWindowHierarchy(topHwnd, contentHwnd)
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

    if !WinExist("ahk_id " hwnd)
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
        hierarchySummary: candidate.hierarchySummary,
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
    record.hierarchySummary := candidate.hierarchySummary
    record.processName := candidate.processName
    record.rootOwner := candidate.rootOwner

    if (record.topHwnd != candidate.topHwnd || record.contentHwnd != candidate.contentHwnd) {
        RebindTrackedTab(host, tabId, candidate)
        return true
    }

    if WinExist("ahk_id " record.contentHwnd) {
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
    LayoutTabButtons(host)
    ShowOnlyActiveTab(host)
    UpdateHostTitle(host)
    RedrawHostWindow(host)
    if host.hwnd && WinExist("ahk_id " host.hwnd)
        WinActivate("ahk_id " host.hwnd)
    ; Destroy empty popout
    if host.isPopout && host.tabOrder.Length = 0 {
        for i, h in g_PopoutHosts {
            if h = host {
                g_PopoutHosts.RemoveAt(i)
                break
            }
        }
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

    if !WinExist("ahk_id " hwnd)
        return false

    ; Grab focus before hiding the original window so Windows doesn't redirect it elsewhere.
    if host.hwnd && WinExist("ahk_id " host.hwnd)
        DllCall("SetForegroundWindow", "ptr", host.hwnd)

    if (record.topHwnd != hwnd) && WinExist("ahk_id " record.topHwnd) {
        record.sourceWasHidden := true
        DllCall("ShowWindow", "ptr", record.topHwnd, "int", 0)
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

    SetWindowLongPtrValue(hwnd, -8, 0)
    SetWindowLongPtrValue(hwnd, -16, newStyle)
    SetWindowLongPtrValue(hwnd, -20, newExStyle)
    DllCall("SetParent", "ptr", hwnd, "ptr", host.clientHwnd, "ptr")

    ; Position at final content rect immediately so when we show it there's no resize glitch.
    GetEmbedRect(host, &areaX, &areaY, &areaW, &areaH)
    areaX += 1
    areaY += 1
    areaW -= 2
    areaH -= 2
    flags := 0x0020 | 0x0040 | 0x0004 | 0x0010
    DllCall("SetWindowPos", "ptr", hwnd, "ptr", 0, "int", areaX, "int", areaY, "int", areaW, "int", areaH, "uint", flags)
    DllCall("ShowWindow", "ptr", hwnd, "int", 0)  ; hide until explicitly shown by ShowOnlyActiveTab
    return true
}

; Reparents content back to original parent, restores style/position; shows top if it was hidden.
DetachTrackedWindow(host, tabId, restoreWindow := true, restoreSource := true) {
    if !host.tabRecords.Has(tabId)
        return

    record := host.tabRecords[tabId]
    hwnd := record.contentHwnd

    ; When closing, claim host focus before any ShowWindow so focus stays in StackTabs.
    if !restoreWindow && host.hwnd && WinExist("ahk_id " host.hwnd)
        DllCall("SetForegroundWindow", "ptr", host.hwnd)

    ; Show parent FIRST (critical for WinUI/XAML apps like PowerShell/Windows Terminal)
    ; so the composition tree can reattach before we reparent the content.
    ; Skip when closing: topHwnd is about to close anyway and showing it steals focus.
    if restoreSource && restoreWindow && record.sourceWasHidden && (record.topHwnd != hwnd) && DllCall("IsWindow", "ptr", record.topHwnd) {
        DllCall("ShowWindow", "ptr", record.topHwnd, "int", record.sourceWasVisible ? 5 : 0)
    }

    if WinExist("ahk_id " hwnd) {
        ; Validate parent: if destroyed, fall back to desktop (top-level window)
        parentHwnd := record.originalContentParent
        if !parentHwnd || !DllCall("IsWindow", "ptr", parentHwnd)
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

        DllCall("SetParent", "ptr", hwnd, "ptr", parentHwnd, "ptr")
        SetWindowLongPtrValue(hwnd, -8, record.originalContentOwner)
        SetWindowLongPtrValue(hwnd, -16, record.originalContentStyle)
        SetWindowLongPtrValue(hwnd, -20, record.originalContentExStyle)

        flags := 0x0020 | 0x0040
        DllCall("SetWindowPos", "ptr", hwnd, "ptr", 0
            , "int", posX, "int", posY
            , "int", record.originalContentW, "int", record.originalContentH
            , "uint", flags)
        DllCall("ShowWindow", "ptr", hwnd, "int", restoreWindow ? 5 : 0)
    }
    record.sourceWasHidden := false
}

; Moves embedded window from source host to dest host (popout/merge).
TransferTrackedWindow(sourceHost, destHost, tabId) {
    if !destHost.tabRecords.Has(tabId)
        return false

    record := destHost.tabRecords[tabId]
    hwnd := record.contentHwnd

    if !WinExist("ahk_id " hwnd)
        return false

    ; Direct reparent: source host's client -> dest host's client
    ; Window stays as WS_CHILD the whole time - no restore to original
    DllCall("SetParent", "ptr", hwnd, "ptr", destHost.clientHwnd, "ptr")

    flags := 0x0020 | 0x0040 | 0x0004 | 0x0010
    DllCall("SetWindowPos", "ptr", hwnd, "ptr", 0, "int", 0, "int", 0, "int", 100, "int", 100, "uint", flags)
    DllCall("ShowWindow", "ptr", hwnd, "int", 4)
    RedrawEmbeddedWindow(hwnd)
    return true
}

; Positions and sizes tab buttons, popout/close controls, and indicators.
LayoutTabButtons(host, windowWidth := 0, windowHeight := 0) {
    global g_HostWidth, g_HostHeight, g_HostPadding, g_TabGap, g_MinTabWidth, g_MaxTabWidth, g_TabHeight
    global g_CloseButtonWidth, g_PopoutButtonWidth, g_TabSlotMax, g_HeaderHeight, g_TabBarAlignment, g_TabBarOffsetY
    global g_UseCustomTitleBar, g_TitleBarHeight, g_TabPosition, g_TabIndicatorHeight

    if !host || !host.gui
        return
    if !host.hwnd || !WinExist("ahk_id " host.hwnd)
        return

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
    extraBtnWidth := g_CloseButtonWidth + g_PopoutButtonWidth  ; popout + close per tab

    if !tabCount {
        for _, ctrl in host.tabSlotButtons
            ctrl.Visible := false
        for _, ctrl in host.tabSlotCloseButtons
            ctrl.Visible := false
        for _, ctrl in host.tabSlotPopoutButtons
            ctrl.Visible := false
        host.tabButtons := Map()
        host.tabCloseButtons := Map()
        host.tabPopoutButtons := Map()
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
    tabBtnY := tabBarY + tabOffsetY
    needed := Min(tabCount, g_TabSlotMax)
    while host.tabSlotButtons.Length < needed {
        btn := host.gui.Add("Text", "Hidden x0 y" tabBtnY " w100 h" g_TabHeight " +0x200 +0x100 Center", "")
        btn.SetFont("s" g_ThemeFontSize, g_ThemeFontNameTab)
        btn.OnEvent("Click", SelectSlot)
        host.tabSlotButtons.Push(btn)
        popoutBtn := host.gui.Add("Text", "Hidden x0 y" tabBtnY " w" g_PopoutButtonWidth " h" g_TabHeight " +0x200 +0x100 Center", "")
        popoutBtn.SetFont("s" g_ThemeIconFontSize, g_ThemeIconFont)
        popoutBtn.Opt("c" g_ThemeIconColor)
        popoutBtn.OnEvent("Click", PopOutSlot)
        host.tabSlotPopoutButtons.Push(popoutBtn)
        closeBtn := host.gui.Add("Text", "Hidden x0 y" tabBtnY " w" g_CloseButtonWidth " h" g_TabHeight " +0x200 +0x100 Center", g_IconClose)
        closeBtn.SetFont("s" g_ThemeIconFontSize, g_ThemeIconFont)
        closeBtn.Opt("c" g_ThemeIconColor)
        closeBtn.OnEvent("Click", CloseSlot)
        host.tabSlotCloseButtons.Push(closeBtn)
        if g_TabIndicatorHeight > 0 {
            global g_ThemeTabIndicatorColor, g_ThemeTabActiveBg
            color := (g_ThemeTabIndicatorColor != "") ? g_ThemeTabIndicatorColor : g_ThemeTabActiveBg
            indic := host.gui.Add("Text", "Hidden x0 y0 w100 h" g_TabIndicatorHeight " Background" color, "")
            host.tabSlotIndicators.Push(indic)
        }
    }
    ; Create indicators when switching from TabIndicatorHeight=0 to >0 (slots exist but no indicators)
    while host.tabSlotIndicators.Length < needed && g_TabIndicatorHeight > 0 {
        global g_ThemeTabIndicatorColor, g_ThemeTabActiveBg
        color := (g_ThemeTabIndicatorColor != "") ? g_ThemeTabIndicatorColor : g_ThemeTabActiveBg
        indic := host.gui.Add("Text", "Hidden x0 y0 w100 h" g_TabIndicatorHeight " Background" color, "")
        host.tabSlotIndicators.Push(indic)
    }

    usableWidth := Max(200, windowWidth - (g_HostPadding * 2))
    tabWidth := Floor((usableWidth - ((tabCount - 1) * g_TabGap)) / tabCount)
    tabWidth := Max(g_MinTabWidth, Min(g_MaxTabWidth, tabWidth))
    titleWidth := tabWidth - extraBtnWidth

    host.tabButtons := Map()
    host.tabCloseButtons := Map()
    host.tabPopoutButtons := Map()
    host.tabIndicators := Map()
    x := g_HostPadding
    for i, tabId in host.tabOrder {
        if i > g_TabSlotMax
            break
        btn := host.tabSlotButtons[i]
        popoutBtn := host.tabSlotPopoutButtons[i]
        closeBtn := host.tabSlotCloseButtons[i]
        title := host.tabRecords.Has(tabId) ? ShortTitle(FilterTitle(host.tabRecords[tabId].title), g_TabTitleMaxLen) : "Window"
        btn.Text := title
        btn.Move(x, tabBtnY, titleWidth, g_TabHeight)
        btn.tabSlotHost := host
        btn.tabSlotId   := tabId
        btn.Visible := true
        host.tabButtons[tabId] := btn

        popoutBtn.Move(x + titleWidth, tabBtnY, g_PopoutButtonWidth, g_TabHeight)
        popoutBtn.tabSlotHost    := host
        popoutBtn.tabSlotId      := tabId
        popoutBtn.tabSlotIsMerge := host.isPopout
        if host.isPopout
            popoutBtn.Text := g_IconMerge
        else
            popoutBtn.Text := g_IconPopout
        popoutBtn.Visible := true
        host.tabPopoutButtons[tabId] := popoutBtn

        closeBtn.Move(x + titleWidth + g_PopoutButtonWidth, tabBtnY, g_CloseButtonWidth, g_TabHeight)
        closeBtn.tabSlotHost := host
        closeBtn.tabSlotId   := tabId
        closeBtn.Visible := true
        host.tabCloseButtons[tabId] := closeBtn

        if g_TabIndicatorHeight > 0 && host.tabSlotIndicators.Length >= i {
            indicY := (g_TabPosition = "bottom") ? tabBtnY : (tabBtnY + g_TabHeight - g_TabIndicatorHeight)
            ; Indicator only under title area to avoid z-order issues with popout/close icons
            host.tabSlotIndicators[i].Move(x, indicY, titleWidth, g_TabIndicatorHeight)
            host.tabIndicators[tabId] := host.tabSlotIndicators[i]
        }

        x += tabWidth + g_TabGap
    }

    Loop Max(0, host.tabSlotButtons.Length - tabCount) {
        i := tabCount + A_Index
        host.tabSlotButtons[i].Visible := false
        host.tabSlotPopoutButtons[i].Visible := false
        host.tabSlotCloseButtons[i].Visible := false
        if host.tabSlotIndicators.Length >= i
            host.tabSlotIndicators[i].Visible := false
    }

    UpdateTabButtonStyles(host)
}

; Tab click handler: selects the tab for this slot.
SelectSlot(ctrl, *) {
    if ctrl.HasProp("tabSlotHost") && ctrl.HasProp("tabSlotId")
        SelectTab(ctrl.tabSlotHost, ctrl.tabSlotId)
}

; Close button handler: closes the tab for this slot.
CloseSlot(ctrl, *) {
    if ctrl.HasProp("tabSlotHost") && ctrl.HasProp("tabSlotId")
        CloseTab(ctrl.tabSlotHost, ctrl.tabSlotId)
}

; Popout/merge button handler: popout to new window or merge back to main.
PopOutSlot(ctrl, *) {
    if ctrl.HasProp("tabSlotHost") && ctrl.HasProp("tabSlotId") && ctrl.HasProp("tabSlotIsMerge") {
        if ctrl.tabSlotIsMerge
            MergeBackTab(ctrl.tabSlotHost, ctrl.tabSlotId)
        else
            PopOutTab(ctrl.tabSlotHost, ctrl.tabSlotId)
    }
}


; Timer: updates tab hover state for all hosts based on mouse position.
CheckTabHoverAll() {
    MouseGetPos(, , , &ctrlHwnd, 2)
    for host in GetAllHosts() {
        if !host.hwnd || !host.HasProp("tabButtons")
            continue
        newHovered := ""
        for tabId, btn in host.tabButtons {
            if btn.Hwnd = ctrlHwnd {
                newHovered := tabId
                break
            }
        }
        if newHovered != host.tabHoveredId {
            host.tabHoveredId := newHovered
            UpdateTabButtonStyles(host)
        }
    }
}

; Sets active tab, shows its content, updates host title.
SelectTab(host, tabId, *) {
    if !host.tabRecords.Has(tabId)
        return

    host.activeTabId := tabId
    ShowOnlyActiveTab(host)
    UpdateHostTitle(host)
    if host.hwnd && WinExist("ahk_id " host.hwnd)
        WinActivate("ahk_id " host.hwnd)
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
    RedrawHostWindow(sourceHost)
    RedrawHostWindow(popoutHost)
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
    popoutHost.gui.Destroy()

    LayoutTabButtons(g_MainHost)
    ShowOnlyActiveTab(g_MainHost)
    UpdateHostTitle(g_MainHost)
    RedrawHostWindow(g_MainHost)
}

; Positions host2 (popout) on the opposite half of the monitor from host1.
ArrangeHostsSideBySide(host1, host2) {
    try {
        WinGetPos(&x1, &y1, &w1, &h1, "ahk_id " host1.hwnd)
        gap := 8

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
        workH := workB - workT
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
        UpdateTabButtonStyles(host)
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
        flags := 0x0020 | 0x0004 | 0x0010
        DllCall("SetWindowPos", "ptr", record.contentHwnd, "ptr", 0
            , "int", areaX, "int", areaY, "int", areaW, "int", areaH
            , "uint", flags)
        DllCall("ShowWindow", "ptr", record.contentHwnd, "int", 4)  ; SW_SHOWNA
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
            DllCall("ShowWindow", "ptr", record.contentHwnd, "int", 0)
    }
    ; Redraw after layout is settled
    if activeHwnd
        RedrawEmbeddedWindow(activeHwnd)

    UpdateTabButtonStyles(host)
    host.lastRefreshActiveTabId := host.activeTabId
    if host.activeTabId != ""
        SetTimer(DeferredRepaintCheck.Bind(host), -50)
}

; Applies active/inactive/hover styles to tab buttons and indicators.
UpdateTabButtonStyles(host) {
    global g_ThemeTabActiveBg, g_ThemeTabActiveText, g_ThemeTabInactiveBg, g_ThemeTabInactiveBgHover
    global g_ThemeTabInactiveText, g_ThemeIconColor, g_ThemeFontName, g_ThemeFontNameTab, g_ThemeFontSize
    global g_ActiveTabStyle, g_TabIndicatorHeight
    if !host || !IsObject(host) || !host.HasProp("tabButtons") || !host.tabButtons
        return
    if !host.hwnd || !WinExist("ahk_id " host.hwnd)
        return
    hoveredId := host.HasProp("tabHoveredId") ? host.tabHoveredId : ""
    for tabId, ctrl in host.tabButtons {
        title := host.tabRecords.Has(tabId) ? ShortTitle(FilterTitle(host.tabRecords[tabId].title), g_TabTitleMaxLen) : "Window"
        if tabId = host.activeTabId {
            ctrl.Text := title
            ctrl.SetFont("s" g_ThemeFontSize " Bold", g_ThemeFontNameTab)
            if (g_ActiveTabStyle = "indicator") {
                ; Indicator-only: same bg as inactive, accent via indicator strip.
                ; Use TabInactiveText (not TabActiveText) for readability â€” TabActiveText is tuned for bright TabActiveBg
                ctrl.Opt("Background0x" g_ThemeTabInactiveBg " c" g_ThemeTabInactiveText)
                if host.tabCloseButtons.Has(tabId)
                    host.tabCloseButtons[tabId].Opt("Background0x" g_ThemeTabInactiveBg " c" g_ThemeIconColor)
                if host.tabPopoutButtons.Has(tabId)
                    host.tabPopoutButtons[tabId].Opt("Background0x" g_ThemeTabInactiveBg " c" g_ThemeIconColor)
            } else {
                ; Full: active tab has distinct background
                ctrl.Opt("Background0x" g_ThemeTabActiveBg " c" g_ThemeTabActiveText)
                if host.tabCloseButtons.Has(tabId)
                    host.tabCloseButtons[tabId].Opt("Background0x" g_ThemeTabActiveBg " c" g_ThemeTabActiveText)
                if host.tabPopoutButtons.Has(tabId)
                    host.tabPopoutButtons[tabId].Opt("Background0x" g_ThemeTabActiveBg " c" g_ThemeTabActiveText)
            }
            if host.tabIndicators.Has(tabId)
                host.tabIndicators[tabId].Visible := (g_TabIndicatorHeight > 0)
        } else {
            inactiveBg := (tabId = hoveredId) ? g_ThemeTabInactiveBgHover : g_ThemeTabInactiveBg
            ctrl.Text := title
            ctrl.SetFont("s" g_ThemeFontSize " Norm", g_ThemeFontName)
            ctrl.Opt("Background0x" inactiveBg " c" g_ThemeTabInactiveText)
            if host.tabCloseButtons.Has(tabId)
                host.tabCloseButtons[tabId].Opt("Background0x" inactiveBg " c" g_ThemeIconColor)
            if host.tabPopoutButtons.Has(tabId)
                host.tabPopoutButtons[tabId].Opt("Background0x" inactiveBg " c" g_ThemeIconColor)
            if host.tabIndicators.Has(tabId)
                host.tabIndicators[tabId].Visible := false
        }
    }
    ; Hide all indicators when TabIndicatorHeight=0 (e.g. after switching from a theme that had them)
    if g_TabIndicatorHeight = 0 {
        for _, indic in host.tabSlotIndicators
            indic.Visible := false
    }
}

; Updates host window title with tab count and active tab name.
UpdateHostTitle(host) {
    global g_HostTitle

    if !host || !host.gui
        return

    liveCount := 0
    for tabId in host.tabOrder {
        if host.tabRecords.Has(tabId) && IsWindowExists(host.tabRecords[tabId].contentHwnd)
            liveCount++
    }
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

    if host.hwnd && WinExist("ahk_id " host.hwnd) {
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

; Exit handler: restores all tabs, saves session, clears state.
CleanupAll(*) {
    global g_MainHost, g_PopoutHosts, g_IsCleaningUp, g_PendingCandidates, g_SessionPath, g_WatchdogTimerActive
    global g_WinEventHooks, g_WinEventHookCallback

    if g_IsCleaningUp
        return

    g_IsCleaningUp := true
    g_WatchdogTimerActive := false
    SetTimer(WatchdogCheck, 0)

    if IsSet(g_WinEventHooks) {
        for hook in g_WinEventHooks
            DllCall("UnhookWinEvent", "Ptr", hook)
        g_WinEventHooks := []
    }

    ; Save main window position/size to session.ini (only if window still exists)
    if IsObject(g_MainHost) && g_MainHost.hwnd && WinExist("ahk_id " g_MainHost.hwnd) {
        try {
            WinGetPos(&wx, &wy, &ww, &wh, "ahk_id " g_MainHost.hwnd)
            IniWrite(wx, g_SessionPath, "Session", "WindowX")
            IniWrite(wy, g_SessionPath, "Session", "WindowY")
            IniWrite(ww, g_SessionPath, "Session", "WindowW")
            IniWrite(wh, g_SessionPath, "Session", "WindowH")
        }
    }

    for host in GetAllHosts() {
        for tabId in host.tabOrder.Clone()
            RemoveTrackedTab(host, tabId, true)
    }

    g_PendingCandidates := Map()
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
    if !WinExist("ahk_id " hwnd)
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

; Forces embedded window to redraw (invalidates and updates).
RedrawEmbeddedWindow(hwnd) {
    if !WinExist("ahk_id " hwnd)
        return

    flags := 0x0001 | 0x0004 | 0x0080 | 0x0100 | 0x0400
    DllCall("RedrawWindow", "ptr", hwnd, "ptr", 0, "ptr", 0, "uint", flags)
    DllCall("UpdateWindow", "ptr", hwnd)
}

; Forces host window to redraw.
RedrawHostWindow(host) {
    if !host || !host.hwnd
        return

    flags := 0x0001 | 0x0004 | 0x0080 | 0x0100 | 0x0400
    DllCall("RedrawWindow", "ptr", host.hwnd, "ptr", 0, "ptr", 0, "uint", flags)
    DllCall("UpdateWindow", "ptr", host.hwnd)
}

; Timer callback: redraws active tab content and host after layout change.
DeferredRepaintCheck(host, *) {
    if !host || !host.hwnd || !WinExist("ahk_id " host.hwnd)
        return
    if host.activeTabId != "" && host.tabRecords.Has(host.activeTabId) {
        record := host.tabRecords[host.activeTabId]
        if IsWindowExists(record.contentHwnd)
            RedrawEmbeddedWindow(record.contentHwnd)
    }
    RedrawHostWindow(host)
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
    if !hIcon {
        DllCall("DeleteObject", "ptr", hBmp)
        DllCall("DeleteObject", "ptr", hMaskBmp)
    }
    return hIcon
}

; Always-on timing log for window detection: writes to winevents.txt (no DebugDiscovery needed).
WinEventLog(text) {
    global g_WinEventLogPath
    try FileAppend(A_Now "." SubStr(A_TickCount, -2) " " text "`n", g_WinEventLogPath, "UTF-8")
}

; Appends text to discovery.txt when DebugDiscovery=1.
AppendDebugLog(text) {
    global g_DebugLogPath, g_DebugDiscovery
    if !g_DebugDiscovery
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
StackTabsHostIsActive() {
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

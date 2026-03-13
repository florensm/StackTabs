; StackTabs - owner-aware embedded ticket host
; AutoHotkey v2

#Requires AutoHotkey v2.0

; ============ CONFIGURATION ============
; Title text that must appear in the popup/shell window.
; Multiple patterns are supported via Match1=/Match2=/... in StackTabs.ini.
g_WindowTitleMatch  := "Ticket details"
g_WindowTitleMatches := []   ; populated from INI; falls back to g_WindowTitleMatch

; Optional EXE filter. Leave blank to match any process.
g_TargetExe := ""

; How often to scan for new / replaced windows.
g_RefreshInterval := 200
g_CaptureDelayMs := 900
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
g_HeaderHeight := 44
g_TabGap := 6
g_MinTabWidth := 120
g_MaxTabWidth := 240
g_TabHeight := 30
g_TabSlotMax := 50
g_CloseButtonWidth := 22
g_PopoutButtonWidth := 22
g_TabBarOffsetY := 7
g_TabPosition   := "top"    ; "top" or "bottom"
g_TabIndicatorHeight := 3  ; height in px of the active-tab indicator strip; 0 to disable

; === THEME defaults (overwritten at startup by the active theme file) ===
g_ThemeBackground      := "1C1C2E"
g_ThemeTabBarBg        := "13132A"
g_ThemeTabActiveBg     := "7B6CF6"
g_ThemeTabActiveText   := "FFFFFF"
g_ThemeTabInactiveBg   := "252540"
g_ThemeTabInactiveBgHover := "30304E"
g_ThemeTabInactiveText := "C5CDF0"
g_ThemeIconColor       := "6878B0"
g_ThemeContentBorder   := "35355A"
g_ThemeWindowText      := "E0E8FF"
g_ThemeFontName        := "Segoe UI"
g_ThemeFontNameTab     := "Segoe UI Semibold"
g_ThemeFontSize        := 9
g_ThemeFontSizeClose   := 10
g_ThemeIconFont        := ""   ; auto-detected at startup; override in theme file with IconFont=
g_ThemeIconFontSize    := 16
g_ActiveThemeFile      := "dark.ini"   ; overridden by ThemeFile= in StackTabs.ini
g_UseCustomTitleBar    := false
g_TitleBarHeight       := 28

; Icon codepoints from Segoe Fluent Icons / Segoe MDL2 Assets (same PUA values)
g_IconClose  := Chr(0xe894)   ; ChromeClose (replace with your choice from the icon list)
g_IconPopout := Chr(0xE8A7)   ; OpenInNewWindow
g_IconMerge  := Chr(0xe944)   ; Back

; === TITLE FILTERS ===
; Strip patterns loaded from [TitleFilters] Strip1/Strip2/... in StackTabs.ini.
; Each is a regex removed from the window title before it appears as a tab label.
g_TitleStripPatterns  := []
; Maximum characters shown in a tab label.
g_TabTitleMaxLen      := 60

; Diagnostics.
g_DebugLogPath := A_ScriptDir "\StackTabs-discovery.txt"

LoadConfigFromIni() {
    iniPath := A_ScriptDir "\StackTabs.ini"
    if !FileExist(iniPath)
        return
    global g_WindowTitleMatch, g_WindowTitleMatches, g_TargetExe, g_RefreshInterval, g_CaptureDelayMs, g_TabDisappearGraceMs
    global g_HostTitle, g_HostWidth, g_HostHeight, g_HostMinWidth, g_HostMinHeight, g_HostX, g_HostY
    global g_HostPadding, g_HeaderHeight, g_TabGap, g_MinTabWidth, g_MaxTabWidth, g_TabHeight
    global g_TabSlotMax, g_CloseButtonWidth, g_PopoutButtonWidth, g_TabBarOffsetY, g_TabPosition
    global g_TitleStripPatterns, g_TabTitleMaxLen
    global g_UseCustomTitleBar, g_TitleBarHeight
    global g_ActiveThemeFile
    try {
        g_WindowTitleMatch := IniRead(iniPath, "General", "WindowTitleMatch", g_WindowTitleMatch)
        g_TargetExe := IniRead(iniPath, "General", "TargetExe", g_TargetExe)
        g_RefreshInterval := Integer(IniRead(iniPath, "General", "RefreshInterval", g_RefreshInterval))
        g_CaptureDelayMs := Integer(IniRead(iniPath, "General", "CaptureDelayMs", g_CaptureDelayMs))
        g_TabDisappearGraceMs := Integer(IniRead(iniPath, "General", "TabDisappearGraceMs", g_TabDisappearGraceMs))
        g_HostTitle := IniRead(iniPath, "Layout", "HostTitle", g_HostTitle)
        g_HostWidth := Integer(IniRead(iniPath, "Layout", "HostWidth", g_HostWidth))
        g_HostHeight := Integer(IniRead(iniPath, "Layout", "HostHeight", g_HostHeight))
        g_HostMinWidth := Integer(IniRead(iniPath, "Layout", "HostMinWidth", g_HostMinWidth))
        g_HostMinHeight := Integer(IniRead(iniPath, "Layout", "HostMinHeight", g_HostMinHeight))
        g_HostPadding := Integer(IniRead(iniPath, "Layout", "HostPadding", g_HostPadding))
        g_HeaderHeight := Integer(IniRead(iniPath, "Layout", "HeaderHeight", g_HeaderHeight))
        g_TabGap := Integer(IniRead(iniPath, "Layout", "TabGap", g_TabGap))
        g_MinTabWidth := Integer(IniRead(iniPath, "Layout", "MinTabWidth", g_MinTabWidth))
        g_MaxTabWidth := Integer(IniRead(iniPath, "Layout", "MaxTabWidth", g_MaxTabWidth))
        g_TabHeight := Integer(IniRead(iniPath, "Layout", "TabHeight", g_TabHeight))
        g_TabSlotMax := Integer(IniRead(iniPath, "Layout", "TabSlotMax", g_TabSlotMax))
        g_CloseButtonWidth := Integer(IniRead(iniPath, "Layout", "CloseButtonWidth", g_CloseButtonWidth))
        g_PopoutButtonWidth := Integer(IniRead(iniPath, "Layout", "PopoutButtonWidth", g_PopoutButtonWidth))
        g_TabBarOffsetY := Integer(IniRead(iniPath, "Layout", "TabBarOffsetY", g_TabBarOffsetY))
        g_TabTitleMaxLen := Integer(IniRead(iniPath, "Layout", "TabTitleMaxLen", g_TabTitleMaxLen))
        g_TabPosition := IniRead(iniPath, "Layout", "TabPosition", "top")
        g_TabIndicatorHeight := Integer(IniRead(iniPath, "Layout", "TabIndicatorHeight", "3"))
        g_UseCustomTitleBar := (IniRead(iniPath, "Layout", "UseCustomTitleBar", g_UseCustomTitleBar ? "1" : "0") = "1")
        g_TitleBarHeight := Integer(IniRead(iniPath, "Layout", "TitleBarHeight", "28"))
        g_ActiveThemeFile := Trim(IniRead(iniPath, "Theme", "ThemeFile", "dark.ini"))
        g_HostX := Integer(IniRead(iniPath, "Session", "WindowX", "-1"))
        g_HostY := Integer(IniRead(iniPath, "Session", "WindowY", "-1"))
        g_HostWidth  := Integer(IniRead(iniPath, "Session", "WindowW", g_HostWidth))
        g_HostHeight := Integer(IniRead(iniPath, "Session", "WindowH", g_HostHeight))
    }
    ; Load multiple match patterns from [General] Match1/Match2/...
    ; Falls back to single WindowTitleMatch if no Match keys are present.
    g_WindowTitleMatches := []
    i := 1
    loop {
        val := IniRead(iniPath, "General", "Match" i, "")
        if val = ""
            break
        g_WindowTitleMatches.Push(val)
        i++
    }
    if g_WindowTitleMatches.Length = 0 && g_WindowTitleMatch != ""
        g_WindowTitleMatches.Push(g_WindowTitleMatch)
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

LoadThemeFromFile(themePath) {
    global g_ThemeBackground, g_ThemeTabBarBg, g_ThemeTabActiveBg, g_ThemeTabActiveText
    global g_ThemeTabInactiveBg, g_ThemeTabInactiveBgHover, g_ThemeTabInactiveText, g_ThemeIconColor
    global g_ThemeContentBorder, g_ThemeWindowText, g_ThemeFontName, g_ThemeFontNameTab
    global g_ThemeFontSize, g_ThemeFontSizeClose, g_ThemeIconFont, g_ThemeIconFontSize
    global g_HostPadding, g_HeaderHeight, g_TabGap, g_MinTabWidth, g_MaxTabWidth, g_TabHeight
    global g_CloseButtonWidth, g_PopoutButtonWidth, g_TabBarOffsetY, g_TabPosition, g_TabIndicatorHeight
    if !FileExist(themePath) {
        MsgBox "Theme file not found:`n" themePath "`n`nFalling back to built-in dark defaults.", "StackTabs", 0x30
        return
    }
    g_ThemeBackground         := IniRead(themePath, "Theme", "Background",           g_ThemeBackground)
    g_ThemeTabBarBg           := IniRead(themePath, "Theme", "TabBarBg",             g_ThemeTabBarBg)
    g_ThemeTabActiveBg        := IniRead(themePath, "Theme", "TabActiveBg",          g_ThemeTabActiveBg)
    g_ThemeTabActiveText      := IniRead(themePath, "Theme", "TabActiveText",        g_ThemeTabActiveText)
    g_ThemeTabInactiveBg      := IniRead(themePath, "Theme", "TabInactiveBg",        g_ThemeTabInactiveBg)
    g_ThemeTabInactiveBgHover := IniRead(themePath, "Theme", "TabInactiveBgHover",   g_ThemeTabInactiveBgHover)
    g_ThemeTabInactiveText    := IniRead(themePath, "Theme", "TabInactiveText",      g_ThemeTabInactiveText)
    g_ThemeIconColor          := IniRead(themePath, "Theme", "IconColor",            g_ThemeIconColor)
    g_ThemeContentBorder      := IniRead(themePath, "Theme", "ContentBorder",        g_ThemeContentBorder)
    g_ThemeWindowText         := IniRead(themePath, "Theme", "WindowText",           g_ThemeWindowText)
    g_ThemeFontName           := IniRead(themePath, "Theme", "FontName",             g_ThemeFontName)
    g_ThemeFontNameTab        := IniRead(themePath, "Theme", "FontNameTab",          g_ThemeFontNameTab)
    g_ThemeFontSize           := Integer(IniRead(themePath, "Theme", "FontSize",     String(g_ThemeFontSize)))
    g_ThemeFontSizeClose      := Integer(IniRead(themePath, "Theme", "FontSizeClose", String(g_ThemeFontSizeClose)))
    g_ThemeIconFont           := IniRead(themePath, "Theme", "IconFont",             g_ThemeIconFont)
    g_ThemeIconFontSize       := Integer(IniRead(themePath, "Theme", "IconFontSize", String(g_ThemeIconFontSize)))
    ; Optional layout overrides — only applied if the theme file includes a [Layout] section
    g_HostPadding        := Integer(IniRead(themePath, "Layout", "HostPadding",        String(g_HostPadding)))
    g_HeaderHeight       := Integer(IniRead(themePath, "Layout", "HeaderHeight",       String(g_HeaderHeight)))
    g_TabGap             := Integer(IniRead(themePath, "Layout", "TabGap",             String(g_TabGap)))
    g_MinTabWidth        := Integer(IniRead(themePath, "Layout", "MinTabWidth",        String(g_MinTabWidth)))
    g_MaxTabWidth        := Integer(IniRead(themePath, "Layout", "MaxTabWidth",        String(g_MaxTabWidth)))
    g_TabHeight          := Integer(IniRead(themePath, "Layout", "TabHeight",          String(g_TabHeight)))
    g_CloseButtonWidth   := Integer(IniRead(themePath, "Layout", "CloseButtonWidth",   String(g_CloseButtonWidth)))
    g_PopoutButtonWidth  := Integer(IniRead(themePath, "Layout", "PopoutButtonWidth",  String(g_PopoutButtonWidth)))
    g_TabBarOffsetY      := Integer(IniRead(themePath, "Layout", "TabBarOffsetY",      String(g_TabBarOffsetY)))
    g_TabIndicatorHeight := Integer(IniRead(themePath, "Layout", "TabIndicatorHeight", String(g_TabIndicatorHeight)))
    g_TabPosition        := IniRead(themePath, "Layout", "TabPosition",                g_TabPosition)
}


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

BuildTrayMenu() {
    global g_ActiveThemeFile
    A_TrayMenu.Delete()
    themeSubMenu := Menu()
    themesDir := A_ScriptDir "\themes"
    if DirExist(themesDir) {
        Loop Files, themesDir "\*.ini" {
            fileName := A_LoopFileName
            displayName := ThemeDisplayName(fileName)
            themeSubMenu.Add(displayName, ThemeMenuHandler.Bind(fileName))
            if (Trim(fileName) = Trim(g_ActiveThemeFile))
                try themeSubMenu.Check(displayName)
        }
    }
    A_TrayMenu.Add("Theme", themeSubMenu)
    A_TrayMenu.Add()
    A_TrayMenu.Add("Exit", (*) => ExitApp())
}

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

ThemeMenuHandler(themeFileName, *) {
    SwitchTheme(themeFileName)
}

SwitchTheme(themeFileName) {
    iniPath := A_ScriptDir "\StackTabs.ini"
    if !FileExist(iniPath) {
        examplePath := A_ScriptDir "\StackTabs.ini.example"
        if FileExist(examplePath)
            FileCopy(examplePath, iniPath)
    }
    IniWrite(themeFileName, iniPath, "Theme", "ThemeFile")
    Reload
}

; ============ STATE ============
g_MainHost := ""             ; HostInstance for main window
g_PopoutHosts := []          ; array of HostInstance for popped-out windows
g_PendingCandidates := Map() ; tabId -> {firstSeen, candidate} (main host only)
g_IsCleaningUp := false

LoadConfigFromIni()
LoadThemeFromFile(A_ScriptDir "\themes\" g_ActiveThemeFile)
DetectIconFont()
BuildTrayMenu()
if g_UseCustomTitleBar
    OnMessage(0x83, OnWmNcCalcSize)
BuildHostInstance(false)  ; create main host
OnExit(CleanupAll)
RefreshWindows()
SetTimer(RefreshWindows, g_RefreshInterval)
SetTimer(CheckTabHoverAll, 50)

; Win+Shift+T toggles the host window.
#+t:: {
    global g_MainHost

    if !g_MainHost
        return

    if WinExist("ahk_id " g_MainHost.hwnd) && WinActive("ahk_id " g_MainHost.hwnd)
        g_MainHost.gui.Hide()
    else
        g_MainHost.gui.Show()
}

; Win+Shift+D writes the current hierarchy scan to disk.
#+d:: {
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

TitleBarDragClick(host, *) {
    MouseGetPos(&mx, &my)
    lParam := (mx & 0xFFFF) | (my << 16)
    PostMessage(0xA1, 2, lParam, , "ahk_id " host.hwnd)
}

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
        CleanupAll()
        ExitApp()
    }
}

IsWindowExists(hwnd) {
    return !!DllCall("IsWindow", "ptr", hwnd, "int")
}

; Close a window reliably (WPF/owned windows may ignore PostMessage; SendMessage waits for processing)
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

GetClientWidth(hwnd) {
    try {
        WinGetClientPos(,, &w,, "ahk_id " hwnd)
        return w
    } catch {
        return 0
    }
}

GetClientHeight(hwnd) {
    try {
        WinGetClientPos(,,, &h, "ahk_id " hwnd)
        return h
    } catch {
        return 0
    }
}

GetAllHosts() {
    global g_MainHost, g_PopoutHosts
    hosts := []
    if g_MainHost
        hosts.Push(g_MainHost)
    for h in g_PopoutHosts
        hosts.Push(h)
    return hosts
}

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

; WM_NCCALCSIZE: extend client area into top non-client region to remove white bar (Windows 10/11)
; OnMessage params: (hwnd, msg, lParam, wParam) per AHK docs
OnWmNcCalcSize(hwnd, msg, lParam, wParam) {
    if !wParam || !lParam  ; wParam=1 means valid rects, lParam=struct pointer
        return
    for host in GetAllHosts() {
        if host.hwnd != hwnd
            continue
        ; Call DefWindowProc first; it modifies rgrc[0] to the client rect
        prevProc := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", -4, "ptr")
        result := DllCall("CallWindowProc", "ptr", prevProc, "ptr", hwnd, "uint", msg, "ptr", wParam, "ptr", lParam, "ptr")
        ; Shrink top border: move client top up by ~7px (Windows 10/11 padded border)
        top := NumGet(lParam, 4, "int")
        NumPut("int", top - 7, lParam, 4)
        return result
    }
}

FindTabByNormalizedTitle(normalizedTitle) {
    if normalizedTitle = ""
        return ""
    for host in GetAllHosts() {
        for tabId, record in host.tabRecords {
            if IsWindowExists(record.contentHwnd) && NormalizeTitle(record.title) = normalizedTitle
                return {host: host, tabId: tabId}
        }
    }
    return ""
}

BuildHostInstance(isPopout := false) {
    global g_MainHost, g_PopoutHosts
    global g_HostTitle, g_HostWidth, g_HostHeight, g_HostMinWidth, g_HostMinHeight
    global g_TabHeight, g_CloseButtonWidth
    global g_UseCustomTitleBar, g_TitleBarHeight

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
    global g_HostX, g_HostY
    if !isPopout && g_HostX >= 0 && g_HostY >= 0
        host.gui.Show("x" g_HostX " y" g_HostY " w" g_HostWidth " h" g_HostHeight)
    else
        host.gui.Show("w" g_HostWidth " h" g_HostHeight)

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
        CleanupAll()
        ExitApp()
    }
}

HostGuiResized(host, guiObj, minMax, width, height) {
    if minMax = -1
        return
    LayoutTabButtons(host, width, height)
    ShowOnlyActiveTab(host)
}

RefreshWindows(*) {
    global g_MainHost, g_IsCleaningUp, g_PendingCandidates, g_CaptureDelayMs, g_TabDisappearGraceMs

    if !g_MainHost || g_IsCleaningUp
        return

    now := A_TickCount

    ; Update all hosts: keep tabs alive, check for stale tabs
    for host in GetAllHosts() {
        if !WinExist("ahk_id " host.hwnd)
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
            candidates := DiscoverCandidateTickets()
            for candidate in candidates {
                currentIds[candidate.id] := true

                if host.tabRecords.Has(candidate.id) {
                    if UpdateTrackedTab(host, candidate.id, candidate)
                        structureChanged := true
                    continue
                }

                if !g_PendingCandidates.Has(candidate.id) {
                    g_PendingCandidates[candidate.id] := {firstSeen: now, candidate: candidate}
                    AppendDebugLog("New candidate`r`n" candidate.hierarchySummary "`r`n")
                    continue
                }

                pending := g_PendingCandidates[candidate.id]
                pending.candidate := candidate
                if (now - pending.firstSeen) >= g_CaptureDelayMs {
                    normalizedTitle := NormalizeTitle(candidate.title)
                    match := (normalizedTitle != "") ? FindTabByNormalizedTitle(normalizedTitle) : ""
                    if match {
                        ; Duplicate ticket: close the extra window, keep existing tab
                        CloseWindowReliably(candidate.topHwnd, candidate.contentHwnd)
                        ; Switch to the existing tab
                        match.host.activeTabId := match.tabId
                        ShowOnlyActiveTab(match.host)
                        UpdateHostTitle(match.host)
                        ; Restore focus: closing owned window gives focus to owner (main App)
                        try WinActivate("ahk_id " match.host.hwnd)
                        if match.host = host
                            structureChanged := true
                    } else if CreateTrackedTab(host, candidate) {
                        host.activeTabId := candidate.id
                        ShowOnlyActiveTab(host)
                        ; Restore focus: embedding steals focus to owner (main App)
                        try WinActivate("ahk_id " host.hwnd)
                        structureChanged := true
                    }
                    g_PendingCandidates.Delete(candidate.id)
                }
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

        ShowOnlyActiveTab(host)
        UpdateHostTitle(host)
    }
}

DiscoverCandidateTickets() {
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
        if g_WindowTitleMatches.Length > 0 {
            matched := false
            for pat in g_WindowTitleMatches {
                if InStr(title, pat, false) {
                    matched := true
                    break
                }
            }
            if !matched
                return ""
        }

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

ScoreContentCandidate(topHwnd, hwnd) {
    global g_WindowTitleMatch

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
    if (title != "") && InStr(title, g_WindowTitleMatch, false)
        score += 500000
    if className = "#32770"
        score -= 250000
    if (className = "Static" || className = "Button")
        score -= 900000

    return score
}

GetDescendantWindows(parentHwnd) {
    result := []
    visited := Map()
    CollectDescendantWindows(parentHwnd, &result, visited)
    return result
}

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

BuildCandidateId(topHwnd, title, processName, contentHwnd) {
    rootOwner := GetRootOwner(topHwnd)
    contentClass := GetWindowClassName(contentHwnd)
    return StrLower(processName) "|" rootOwner "|" NormalizeTitle(title) "|" contentClass
}

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

BuildTrackedRecord(candidate) {
    WinGetPos(&x, &y, &w, &h, "ahk_id " candidate.contentHwnd)

    return {
        id: candidate.id,
        title: candidate.title,
        topHwnd: candidate.topHwnd,
        contentHwnd: candidate.contentHwnd,
        processName: candidate.processName,
        rootOwner: candidate.rootOwner,
        hierarchySummary: candidate.hierarchySummary,
        originalContentParent: DllCall("GetParent", "ptr", candidate.contentHwnd, "ptr"),
        originalContentOwner: GetWindowLongPtrValue(candidate.contentHwnd, -8),
        originalContentStyle: GetWindowLongPtrValue(candidate.contentHwnd, -16),
        originalContentExStyle: GetWindowLongPtrValue(candidate.contentHwnd, -20),
        originalContentX: x,
        originalContentY: y,
        originalContentW: w,
        originalContentH: h,
        sourceWasHidden: false,
        sourceWasVisible: (candidate.topHwnd != candidate.contentHwnd) && DllCall("IsWindowVisible", "ptr", candidate.topHwnd) ? 1 : 0,
        lastSeenTick: A_TickCount
    }
}

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

RebindTrackedTab(host, tabId, candidate) {
    if !host.tabRecords.Has(tabId)
        return

    DetachTrackedWindow(host, tabId, false, false)

    record := BuildTrackedRecord(candidate)
    host.tabRecords[tabId] := record
    AttachTrackedWindow(host, tabId)
    AppendDebugLog("Rebound tab: " tabId "`r`n" candidate.hierarchySummary "`r`n")
}

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

ActivateHostAfterClose(host, *) {
    if host.tabOrder.Length > 0 && WinExist("ahk_id " host.hwnd)
        try WinActivate("ahk_id " host.hwnd)
}

CloseTabDeferredUpdate(host, *) {
    global g_PopoutHosts
    LayoutTabButtons(host)
    ShowOnlyActiveTab(host)
    UpdateHostTitle(host)
    RedrawHostWindow(host)
    ; Restore focus: closing owned window gives focus to owner (main App).
    ; PostMessage is async—close may complete after this runs; re-activate once more.
    if host.tabOrder.Length > 0 {
        try WinActivate("ahk_id " host.hwnd)
        SetTimer(ActivateHostAfterClose.Bind(host), -100)
    }
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

CloseActiveTab(host) {
    if host.activeTabId != ""
        CloseTab(host, host.activeTabId)
}

RemoveTrackedTab(host, tabId, restoreWindow := true) {
    if !host.tabRecords.Has(tabId)
        return

    DetachTrackedWindow(host, tabId, restoreWindow, true)

    for idx, currentId in host.tabOrder {
        if currentId = tabId {
            host.tabOrder.RemoveAt(idx)
            break
        }
    }

    host.tabRecords.Delete(tabId)

    if host.activeTabId = tabId
        host.activeTabId := host.tabOrder.Length ? host.tabOrder[1] : ""
}

AttachTrackedWindow(host, tabId) {
    if !host.tabRecords.Has(tabId)
        return false

    record := host.tabRecords[tabId]
    hwnd := record.contentHwnd

    if !WinExist("ahk_id " hwnd)
        return false

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

    flags := 0x0020 | 0x0040 | 0x0004 | 0x0010
    DllCall("SetWindowPos", "ptr", hwnd, "ptr", 0, "int", 0, "int", 0, "int", 100, "int", 100, "uint", flags)
    DllCall("ShowWindow", "ptr", hwnd, "int", 0)
    RedrawEmbeddedWindow(hwnd)
    return true
}

DetachTrackedWindow(host, tabId, restoreWindow := true, restoreSource := true) {
    if !host.tabRecords.Has(tabId)
        return

    record := host.tabRecords[tabId]
    hwnd := record.contentHwnd

    if WinExist("ahk_id " hwnd) {
        DllCall("SetParent", "ptr", hwnd, "ptr", record.originalContentParent, "ptr")
        SetWindowLongPtrValue(hwnd, -8, record.originalContentOwner)
        SetWindowLongPtrValue(hwnd, -16, record.originalContentStyle)
        SetWindowLongPtrValue(hwnd, -20, record.originalContentExStyle)

        flags := 0x0020 | 0x0040
        DllCall("SetWindowPos", "ptr", hwnd, "ptr", 0
            , "int", record.originalContentX, "int", record.originalContentY
            , "int", record.originalContentW, "int", record.originalContentH
            , "uint", flags)
        DllCall("ShowWindow", "ptr", hwnd, "int", restoreWindow ? 5 : 0)
    }

    if restoreSource && record.sourceWasHidden && (record.topHwnd != hwnd) && WinExist("ahk_id " record.topHwnd) {
        DllCall("ShowWindow", "ptr", record.topHwnd, "int", record.sourceWasVisible ? 5 : 0)
    }
    record.sourceWasHidden := false
}

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

LayoutTabButtons(host, windowWidth := 0, windowHeight := 0) {
    global g_HostWidth, g_HostHeight, g_HostPadding, g_TabGap, g_MinTabWidth, g_MaxTabWidth, g_TabHeight
    global g_CloseButtonWidth, g_PopoutButtonWidth, g_TabSlotMax, g_HeaderHeight, g_TabBarOffsetY
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

    global g_ThemeFontName, g_ThemeFontNameTab, g_ThemeFontSize, g_ThemeFontSizeClose, g_ThemeIconColor
    global g_ThemeIconFont, g_ThemeIconFontSize, g_IconClose, g_IconPopout, g_IconMerge
    tabBtnY := tabBarY + g_TabBarOffsetY
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
            indic := host.gui.Add("Text", "Hidden x0 y0 w100 h" g_TabIndicatorHeight " Background" g_ThemeTabActiveBg, "")
            host.tabSlotIndicators.Push(indic)
        }
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
            host.tabSlotIndicators[i].Move(x, indicY, tabWidth, g_TabIndicatorHeight)
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

; Generic slot handlers — registered once per slot at creation; tabSlot* props updated each layout pass.
SelectSlot(ctrl, *) {
    if ctrl.HasProp("tabSlotHost") && ctrl.HasProp("tabSlotId")
        SelectTab(ctrl.tabSlotHost, ctrl.tabSlotId)
}

CloseSlot(ctrl, *) {
    if ctrl.HasProp("tabSlotHost") && ctrl.HasProp("tabSlotId")
        CloseTab(ctrl.tabSlotHost, ctrl.tabSlotId)
}

PopOutSlot(ctrl, *) {
    if ctrl.HasProp("tabSlotHost") && ctrl.HasProp("tabSlotId") && ctrl.HasProp("tabSlotIsMerge") {
        if ctrl.tabSlotIsMerge
            MergeBackTab(ctrl.tabSlotHost, ctrl.tabSlotId)
        else
            PopOutTab(ctrl.tabSlotHost, ctrl.tabSlotId)
    }
}


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

SelectTab(host, tabId, *) {
    if !host.tabRecords.Has(tabId)
        return

    host.activeTabId := tabId
    ShowOnlyActiveTab(host)
    UpdateHostTitle(host)
}

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
    for idx, currentId in sourceHost.tabOrder {
        if currentId = tabId {
            sourceHost.tabOrder.RemoveAt(idx)
            break
        }
    }
    sourceHost.tabRecords.Delete(tabId)
    if sourceHost.activeTabId = tabId
        sourceHost.activeTabId := sourceHost.tabOrder.Length ? sourceHost.tabOrder[1] : ""

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

ArrangeHostsSideBySide(host1, host2) {
    try {
        WinGetPos(&x1, &y1, &w1, &h1, "ahk_id " host1.hwnd)
        gap := 8
        x2 := x1 + w1 + gap
        y2 := y1

        ; Clamp to the work area of the monitor containing host1
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
        ; Prefer right of host1; if it doesn't fit, try left; then clamp
        if x2 + w1 > workR
            x2 := x1 - w1 - gap
        x2 := Max(workL, Min(x2, workR - w1))
        y2 := Max(workT, Min(y2, workB - h1))

        host2.gui.Show("x" x2 " y" y2 " w" w1 " h" h1)
    } catch {
        host2.gui.Show()
    }
}

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

    for tabId in host.tabOrder {
        if !host.tabRecords.Has(tabId)
            continue

        record := host.tabRecords[tabId]
        if !IsWindowExists(record.contentHwnd)
            continue

        if tabId = host.activeTabId {
            flags := 0x0020 | 0x0004 | 0x0010
            DllCall("SetWindowPos", "ptr", record.contentHwnd, "ptr", 0
                , "int", areaX, "int", areaY, "int", areaW, "int", areaH
                , "uint", flags)
            DllCall("ShowWindow", "ptr", record.contentHwnd, "int", 4)
            RedrawEmbeddedWindow(record.contentHwnd)
        } else {
            DllCall("ShowWindow", "ptr", record.contentHwnd, "int", 0)
        }
    }

    UpdateTabButtonStyles(host)
}

UpdateTabButtonStyles(host) {
    global g_ThemeTabActiveBg, g_ThemeTabActiveText, g_ThemeTabInactiveBg, g_ThemeTabInactiveBgHover
    global g_ThemeTabInactiveText, g_ThemeIconColor, g_ThemeFontName, g_ThemeFontNameTab, g_ThemeFontSize
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
            ctrl.Opt("Background0x" g_ThemeTabActiveBg " c" g_ThemeTabActiveText)
            if host.tabCloseButtons.Has(tabId)
                host.tabCloseButtons[tabId].Opt("Background0x" g_ThemeTabActiveBg " c" g_ThemeTabActiveText)
            if host.tabPopoutButtons.Has(tabId)
                host.tabPopoutButtons[tabId].Opt("Background0x" g_ThemeTabActiveBg " c" g_ThemeTabActiveText)
            if host.tabIndicators.Has(tabId)
                host.tabIndicators[tabId].Visible := true
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
}

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
    host.gui.Title := title
    if host.HasProp("titleText") && host.titleText
        host.titleText.Text := title
    UpdateHostIcon(host)
}

GetEmbedRect(host, &x, &y, &w, &h) {
    global g_HostWidth, g_HostHeight, g_HostPadding, g_HeaderHeight
    global g_UseCustomTitleBar, g_TitleBarHeight, g_TabPosition

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
                h := Max(140, clientH - y - g_HeaderHeight - g_HostPadding)
            else
                h := Max(140, clientH - y - g_HostPadding)
            return
        }
    }

    w := Max(200, g_HostWidth - (g_HostPadding * 2))
    if g_TabPosition = "bottom"
        h := Max(140, g_HostHeight - y - g_HeaderHeight - g_HostPadding)
    else
        h := Max(140, g_HostHeight - y - g_HostPadding)
}

CleanupAll(*) {
    global g_MainHost, g_PopoutHosts, g_IsCleaningUp, g_PendingCandidates

    if g_IsCleaningUp
        return

    g_IsCleaningUp := true

    ; Save main window position/size to [Session] in StackTabs.ini
    if IsObject(g_MainHost) && g_MainHost.hwnd {
        iniPath := A_ScriptDir "\StackTabs.ini"
        if !FileExist(iniPath) && FileExist(A_ScriptDir "\StackTabs.ini.example")
            FileCopy(A_ScriptDir "\StackTabs.ini.example", iniPath)
        if FileExist(iniPath) {
            WinGetPos(&wx, &wy, &ww, &wh, "ahk_id " g_MainHost.hwnd)
            IniWrite(wx, iniPath, "Session", "WindowX")
            IniWrite(wy, iniPath, "Session", "WindowY")
            IniWrite(ww, iniPath, "Session", "WindowW")
            IniWrite(wh, iniPath, "Session", "WindowH")
        }
    }

    for host in GetAllHosts() {
        for tabId in host.tabOrder.Clone()
            RemoveTrackedTab(host, tabId, true)
    }

    g_PendingCandidates := Map()
    g_IsCleaningUp := false
}

DumpDiscoveryDebug() {
    global g_DebugLogPath

    discovered := DiscoverCandidateTickets()
    text := "Timestamp: " FormatTime(, "yyyy-MM-dd HH:mm:ss") "`r`n"
    text .= "Discovered tickets: " discovered.Length "`r`n`r`n"

    for candidate in discovered {
        text .= "Tab ID: " candidate.id "`r`n"
        text .= candidate.hierarchySummary "`r`n"
        text .= "------------------------------`r`n"
    }

    FileDelete(g_DebugLogPath)
    FileAppend(text, g_DebugLogPath, "UTF-8")
    MsgBox("Wrote discovery info to:`n" g_DebugLogPath, "StackTabs Debug")
}

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

GetPreferredTabTitle(record) {
    title := SafeWinGetTitle(record.topHwnd)
    if title != ""
        return title
    return SafeWinGetTitle(record.contentHwnd)
}

RedrawEmbeddedWindow(hwnd) {
    if !WinExist("ahk_id " hwnd)
        return

    flags := 0x0001 | 0x0004 | 0x0080 | 0x0100 | 0x0400
    DllCall("RedrawWindow", "ptr", hwnd, "ptr", 0, "ptr", 0, "uint", flags)
}

RedrawHostWindow(host) {
    if !host || !host.hwnd
        return

    flags := 0x0001 | 0x0004 | 0x0080 | 0x0100 | 0x0400
    DllCall("RedrawWindow", "ptr", host.hwnd, "ptr", 0, "ptr", 0, "uint", flags)
}

; ============ ICON ============

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
    SendMessage(0x0080, 0, hBadged,, "ahk_id " host.hwnd)  ; WM_SETICON ICON_SMALL
    SendMessage(0x0080, 1, hBadged,, "ahk_id " host.hwnd)  ; WM_SETICON ICON_BIG
}

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

; Draws hSourceIcon onto a 32x32 bitmap and overlays a coloured dot badge
; in the bottom-right corner.  Returns an HICON the caller owns (DestroyIcon when done).
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

    ; Monochrome AND-mask — all black = fully opaque
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

AppendDebugLog(text) {
    global g_DebugLogPath
    FileAppend("[" FormatTime(, "yyyy-MM-dd HH:mm:ss") "]`r`n" text, g_DebugLogPath, "UTF-8")
}

JoinLines(lines) {
    text := ""
    for idx, line in lines {
        if idx > 1
            text .= "`r`n"
        text .= line
    }
    return text
}

NormalizeTitle(title) {
    normalized := Trim(StrLower(title))
    normalized := RegExReplace(normalized, "\s+", " ")
    return normalized
}

FilterTitle(title) {
    global g_TitleStripPatterns
    for pattern in g_TitleStripPatterns
        title := RegExReplace(title, pattern, "")
    return Trim(title)
}

ShortTitle(title, maxLen := 28) {
    if StrLen(title) <= maxLen
        return title
    return SubStr(title, 1, maxLen - 1) . "..."
}

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

StackTabsHostIsActive() {
    return !!GetActiveStackTabsHost()
}

GetActiveStackTabsHost() {
    try
        activeHwnd := WinGetID("A")
    catch
        return ""
    if !activeHwnd
        return ""
    return GetHostForHwnd(activeHwnd)
}

SafeWinGetTitle(hwnd) {
    try return WinGetTitle("ahk_id " hwnd)
    catch
        return ""
}

SafeWinGetProcessName(hwnd) {
    try return WinGetProcessName("ahk_id " hwnd)
    catch
        return ""
}

GetWindowClassName(hwnd) {
    buf := Buffer(512, 0)
    DllCall("GetClassName", "ptr", hwnd, "ptr", buf, "int", 256)
    return StrGet(buf)
}

GetRootOwner(hwnd) {
    return DllCall("GetAncestor", "ptr", hwnd, "uint", 3, "ptr")
}

GetWindowOwner(hwnd) {
    return DllCall("GetWindow", "ptr", hwnd, "uint", 4, "ptr")
}

GetWindowLongPtrValue(hwnd, index) {
    if A_PtrSize = 8
        return DllCall("GetWindowLongPtr", "ptr", hwnd, "int", index, "ptr")
    return DllCall("GetWindowLong", "ptr", hwnd, "int", index, "ptr")
}

SetWindowLongPtrValue(hwnd, index, value) {
    if A_PtrSize = 8
        return DllCall("SetWindowLongPtr", "ptr", hwnd, "int", index, "ptr", value, "ptr")
    return DllCall("SetWindowLong", "ptr", hwnd, "int", index, "ptr", value, "ptr")
}

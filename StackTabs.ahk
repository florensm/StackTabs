; StackTabs - owner-aware embedded ticket host
; AutoHotkey v2

#Requires AutoHotkey v2.0

; ============ CONFIGURATION ============
; Title text that must appear in the popup/shell window.
g_WindowTitleMatch := "Brave"

; Optional EXE filter. Leave blank to match any process.
g_TargetExe := ""

; How often to scan for new / replaced windows.
g_RefreshInterval := 500
g_CaptureDelayMs := 900
g_TabDisappearGraceMs := 2500

; Host window defaults.
g_HostTitle := "StackTabs"
g_HostWidth := 1200
g_HostHeight := 800
g_HostMinWidth := 700
g_HostMinHeight := 500
g_HostPadding := 8
g_HeaderHeight := 44
g_TabGap := 6
g_MinTabWidth := 120
g_MaxTabWidth := 240
g_TabHeight := 30

; Diagnostics.
g_DebugLogPath := A_ScriptDir "\StackTabs-discovery.txt"

; ============ STATE ============
g_HostGui := ""
g_HostHwnd := 0
g_ClientHwnd := 0
g_StatusText := ""
g_TabButtons := Map()        ; tabId -> Gui control
g_TabRecords := Map()        ; tabId -> record object
g_TabOrder := []             ; ordered tab ids
g_HwndToTabId := Map()       ; hwnd -> tabId
g_PendingCandidates := Map() ; tabId -> {firstSeen, candidate}
g_ActiveTabId := ""
g_IsCleaningUp := false

BuildHostGui()
OnExit(CleanupAll)
RefreshWindows()
SetTimer(RefreshWindows, g_RefreshInterval)

; Win+Shift+T toggles the host window.
#+t:: {
    global g_HostGui

    if !g_HostGui
        return

    if WinExist("ahk_id " g_HostGui.Hwnd) && WinActive("ahk_id " g_HostGui.Hwnd)
        g_HostGui.Hide()
    else
        g_HostGui.Show()
}

; Win+Shift+D writes the current hierarchy scan to disk.
#+d:: {
    DumpDiscoveryDebug()
}

#HotIf StackTabsHostIsActive()
^Tab:: {
    CycleTabs(1)
}
^+Tab:: {
    CycleTabs(-1)
}
#HotIf

BuildHostGui() {
    global g_HostGui, g_HostHwnd, g_ClientHwnd, g_StatusText
    global g_HostTitle, g_HostWidth, g_HostHeight, g_HostMinWidth, g_HostMinHeight
    global g_WindowTitleMatch

    if g_HostGui
        g_HostGui.Destroy()

    g_HostGui := Gui("+Resize +MinSize" g_HostMinWidth "x" g_HostMinHeight, g_HostTitle)
    g_HostGui.BackColor := "1E1E1E"
    g_HostGui.MarginX := 0
    g_HostGui.MarginY := 0
    g_HostGui.SetFont("s10 cWhite", "Segoe UI")
    g_HostGui.OnEvent("Close", HostGuiClosed)

    ; Keep a hidden status control so existing update helpers stay simple.
    g_StatusText := g_HostGui.Add("Text", "Hidden x0 y0 w0 h0", "")

    g_HostGui.Show("w" g_HostWidth " h" g_HostHeight)
    g_HostHwnd := g_HostGui.Hwnd
    g_ClientHwnd := g_HostHwnd
    g_HostGui.OnEvent("Size", HostGuiResized)
}

HostGuiClosed(*) {
    CleanupAll()
    ExitApp()
}

HostGuiResized(guiObj, minMax, width, height) {
    if minMax = -1
        return

    LayoutTabButtons(width)
    ShowOnlyActiveTab()
}

RefreshWindows(*) {
    global g_HostHwnd, g_ClientHwnd, g_IsCleaningUp, g_TabOrder, g_TabRecords, g_ActiveTabId
    global g_PendingCandidates, g_CaptureDelayMs, g_TabDisappearGraceMs

    if !g_ClientHwnd || !g_HostHwnd || g_IsCleaningUp
        return
    if !WinExist("ahk_id " g_HostHwnd)
        return

    now := A_TickCount
    structureChanged := false
    currentIds := Map()

    ; Keep existing embedded tabs alive even when their wrapper window is hidden by us.
    for tabId in g_TabOrder {
        if !g_TabRecords.Has(tabId)
            continue

        record := g_TabRecords[tabId]
        if WinExist("ahk_id " record.contentHwnd) {
            record.lastSeenTick := now
            title := GetPreferredTabTitle(record)
            if title != ""
                record.title := title
        }
    }

    candidates := DiscoverCandidateTickets()
    for candidate in candidates {
        currentIds[candidate.id] := true

        if g_TabRecords.Has(candidate.id) {
            if UpdateTrackedTab(candidate.id, candidate)
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
            if CreateTrackedTab(candidate)
                structureChanged := true
            g_PendingCandidates.Delete(candidate.id)
        }
    }

    stalePending := []
    for tabId, pending in g_PendingCandidates {
        if !currentIds.Has(tabId)
            stalePending.Push(tabId)
    }
    for tabId in stalePending
        g_PendingCandidates.Delete(tabId)

    staleTabs := []
    for tabId in g_TabOrder {
        if !g_TabRecords.Has(tabId)
            continue

        record := g_TabRecords[tabId]
        if WinExist("ahk_id " record.contentHwnd)
            continue
        if (now - record.lastSeenTick) > g_TabDisappearGraceMs
            staleTabs.Push(tabId)
    }

    for tabId in staleTabs {
        RemoveTrackedTab(tabId, false)
        structureChanged := true
    }

    if g_ActiveTabId && !g_TabRecords.Has(g_ActiveTabId)
        g_ActiveTabId := ""
    if (g_ActiveTabId = "") && g_TabOrder.Length
        g_ActiveTabId := g_TabOrder[1]

    if structureChanged
        LayoutTabButtons()

    ShowOnlyActiveTab()
    UpdateStatusText()
    UpdateHostTitle()
}

DiscoverCandidateTickets() {
    global g_HostHwnd, g_TabRecords

    candidates := []
    seenIds := Map()

    for hwnd in WinGetList() {
        if hwnd = g_HostHwnd
            continue

        candidate := BuildCandidateFromTopWindow(hwnd)
        if !IsObject(candidate)
            continue
        if seenIds.Has(candidate.id)
            continue

        seenIds[candidate.id] := true
        candidates.Push(candidate)
    }

    return candidates
}

BuildCandidateFromTopWindow(topHwnd) {
    global g_WindowTitleMatch, g_TargetExe

    try {
        if !WinExist("ahk_id " topHwnd)
            return ""
        if !DllCall("IsWindowVisible", "ptr", topHwnd)
            return ""

        title := WinGetTitle("ahk_id " topHwnd)
        if (title = "")
            return ""
        if (g_WindowTitleMatch != "") && !InStr(title, g_WindowTitleMatch, false)
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

CreateTrackedTab(candidate) {
    global g_TabRecords, g_TabOrder, g_ActiveTabId

    if g_TabRecords.Has(candidate.id)
        return false

    record := BuildTrackedRecord(candidate)
    g_TabRecords[candidate.id] := record

    if !AttachTrackedWindow(candidate.id) {
        g_TabRecords.Delete(candidate.id)
        return false
    }

    IndexTrackedHwnds(candidate.id)
    g_TabOrder.Push(candidate.id)
    if (g_ActiveTabId = "")
        g_ActiveTabId := candidate.id
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

UpdateTrackedTab(tabId, candidate) {
    global g_TabRecords, g_ClientHwnd

    record := g_TabRecords[tabId]
    record.lastSeenTick := A_TickCount
    record.title := candidate.title
    record.hierarchySummary := candidate.hierarchySummary
    record.processName := candidate.processName
    record.rootOwner := candidate.rootOwner

    if (record.topHwnd != candidate.topHwnd || record.contentHwnd != candidate.contentHwnd) {
        RebindTrackedTab(tabId, candidate)
        return true
    }

    if WinExist("ahk_id " record.contentHwnd) {
        currentParent := DllCall("GetParent", "ptr", record.contentHwnd, "ptr")
        if currentParent != g_ClientHwnd {
            AttachTrackedWindow(tabId)
            return true
        }
    }

    return false
}

RebindTrackedTab(tabId, candidate) {
    global g_TabRecords

    if !g_TabRecords.Has(tabId)
        return

    DetachTrackedWindow(tabId, false, false)
    UnindexTrackedHwnds(tabId)

    record := BuildTrackedRecord(candidate)
    g_TabRecords[tabId] := record
    AttachTrackedWindow(tabId)
    IndexTrackedHwnds(tabId)
    AppendDebugLog("Rebound tab: " tabId "`r`n" candidate.hierarchySummary "`r`n")
}

RemoveTrackedTab(tabId, restoreWindow := true) {
    global g_TabRecords, g_TabOrder, g_ActiveTabId

    if !g_TabRecords.Has(tabId)
        return

    DetachTrackedWindow(tabId, restoreWindow, true)
    UnindexTrackedHwnds(tabId)

    for idx, currentId in g_TabOrder {
        if currentId = tabId {
            g_TabOrder.RemoveAt(idx)
            break
        }
    }

    g_TabRecords.Delete(tabId)

    if g_ActiveTabId = tabId
        g_ActiveTabId := g_TabOrder.Length ? g_TabOrder[1] : ""
}

AttachTrackedWindow(tabId) {
    global g_TabRecords, g_ClientHwnd

    if !g_TabRecords.Has(tabId)
        return false

    record := g_TabRecords[tabId]
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
    DllCall("SetParent", "ptr", hwnd, "ptr", g_ClientHwnd, "ptr")

    flags := 0x0020 | 0x0040 | 0x0004 | 0x0010
    DllCall("SetWindowPos", "ptr", hwnd, "ptr", 0, "int", 0, "int", 0, "int", 100, "int", 100, "uint", flags)
    DllCall("ShowWindow", "ptr", hwnd, "int", 0)
    RedrawEmbeddedWindow(hwnd)
    return true
}

DetachTrackedWindow(tabId, restoreWindow := true, restoreSource := true) {
    global g_TabRecords

    if !g_TabRecords.Has(tabId)
        return

    record := g_TabRecords[tabId]
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

IndexTrackedHwnds(tabId) {
    global g_TabRecords, g_HwndToTabId

    if !g_TabRecords.Has(tabId)
        return

    record := g_TabRecords[tabId]
    g_HwndToTabId[record.contentHwnd] := tabId
    if record.topHwnd != record.contentHwnd
        g_HwndToTabId[record.topHwnd] := tabId
}

UnindexTrackedHwnds(tabId) {
    global g_TabRecords, g_HwndToTabId

    if !g_TabRecords.Has(tabId)
        return

    record := g_TabRecords[tabId]
    TryDeleteMapKey(g_HwndToTabId, record.contentHwnd)
    if record.topHwnd != record.contentHwnd
        TryDeleteMapKey(g_HwndToTabId, record.topHwnd)
}

LayoutTabButtons(windowWidth := 0) {
    global g_HostGui, g_HostHwnd, g_HostWidth, g_HostPadding, g_TabButtons
    global g_TabOrder, g_TabRecords, g_ActiveTabId, g_TabGap, g_MinTabWidth, g_MaxTabWidth, g_TabHeight

    if !g_HostGui
        return
    if !g_HostHwnd || !WinExist("ahk_id " g_HostHwnd)
        return

    if !windowWidth {
        try {
            clientX := 0
            clientY := 0
            clientW := 0
            clientH := 0
            WinGetClientPos(&clientX, &clientY, &clientW, &clientH, "ahk_id " g_HostHwnd)
            windowWidth := clientW
        } catch {
            windowWidth := g_HostWidth
        }
    }
    if !windowWidth
        windowWidth := g_HostWidth

    for _, ctrl in g_TabButtons
        try ctrl.Destroy()
    g_TabButtons := Map()

    tabCount := g_TabOrder.Length
    if !tabCount
        return

    usableWidth := Max(200, windowWidth - (g_HostPadding * 2))
    tabWidth := Floor((usableWidth - ((tabCount - 1) * g_TabGap)) / tabCount)
    tabWidth := Max(g_MinTabWidth, Min(g_MaxTabWidth, tabWidth))

    x := g_HostPadding
    for tabId in g_TabOrder {
        title := g_TabRecords.Has(tabId) ? g_TabRecords[tabId].title : "Window"
        btn := g_HostGui.Add("Text", "x" x " y7 w" tabWidth " h" g_TabHeight " +0x200 +0x100 Border Center", ShortTitle(title, 26))
        btn.SetFont("s9", "Segoe UI Semibold")
        btn.OnEvent("Click", SelectTab.Bind(tabId))
        g_TabButtons[tabId] := btn
        x += tabWidth + g_TabGap
    }

    UpdateTabButtonStyles()
}

SelectTab(tabId, *) {
    global g_ActiveTabId, g_TabRecords

    if !g_TabRecords.Has(tabId)
        return

    g_ActiveTabId := tabId
    ShowOnlyActiveTab()
    UpdateStatusText()
    UpdateHostTitle()
}

ShowOnlyActiveTab() {
    global g_TabOrder, g_TabRecords, g_ActiveTabId

    if (g_ActiveTabId != "") && !g_TabRecords.Has(g_ActiveTabId)
        g_ActiveTabId := ""
    if (g_ActiveTabId = "") && g_TabOrder.Length
        g_ActiveTabId := g_TabOrder[1]

    if g_ActiveTabId = "" {
        UpdateTabButtonStyles()
        return
    }

    GetEmbedRect(&areaX, &areaY, &areaW, &areaH)

    for tabId in g_TabOrder {
        if !g_TabRecords.Has(tabId)
            continue

        record := g_TabRecords[tabId]
        if !WinExist("ahk_id " record.contentHwnd)
            continue

        if tabId = g_ActiveTabId {
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

    UpdateTabButtonStyles()
}

UpdateTabButtonStyles() {
    global g_TabButtons, g_ActiveTabId, g_TabRecords

    for tabId, ctrl in g_TabButtons {
        title := g_TabRecords.Has(tabId) ? ShortTitle(g_TabRecords[tabId].title, 26) : "Window"
        if tabId = g_ActiveTabId {
            ctrl.Text := title
            ctrl.SetFont("s9 Bold", "Segoe UI Semibold")
            ctrl.Opt("Background0x2D7DFF cFFFFFF")
        } else {
            ctrl.Text := title
            ctrl.SetFont("s9 Norm", "Segoe UI")
            ctrl.Opt("Background0x30343B cD8DEE9")
        }
    }
}

UpdateStatusText() {
    return
}

UpdateHostTitle() {
    global g_HostGui, g_HostTitle, g_TabOrder, g_ActiveTabId, g_TabRecords

    if !g_HostGui
        return

    if (g_ActiveTabId != "") && g_TabRecords.Has(g_ActiveTabId)
        g_HostGui.Title := g_HostTitle . " (" g_TabOrder.Length ") - " . g_TabRecords[g_ActiveTabId].title
    else
        g_HostGui.Title := g_HostTitle . " (" g_TabOrder.Length ")"
}

GetEmbedRect(&x, &y, &w, &h) {
    global g_HostHwnd, g_HostWidth, g_HostHeight, g_HostPadding, g_HeaderHeight

    x := g_HostPadding
    y := g_HeaderHeight + g_HostPadding

    if g_HostHwnd && WinExist("ahk_id " g_HostHwnd) {
        try {
            clientX := 0
            clientY := 0
            clientW := 0
            clientH := 0
            WinGetClientPos(&clientX, &clientY, &clientW, &clientH, "ahk_id " g_HostHwnd)
            w := Max(200, clientW - (g_HostPadding * 2))
            h := Max(140, clientH - y - g_HostPadding)
            return
        }
    }

    w := Max(200, g_HostWidth - (g_HostPadding * 2))
    h := Max(140, g_HostHeight - y - g_HostPadding)
}

CleanupAll(*) {
    global g_TabOrder, g_IsCleaningUp, g_PendingCandidates, g_HwndToTabId

    if g_IsCleaningUp
        return

    g_IsCleaningUp := true

    restoreList := []
    for tabId in g_TabOrder
        restoreList.Push(tabId)

    for tabId in restoreList
        RemoveTrackedTab(tabId, true)

    g_PendingCandidates := Map()
    g_HwndToTabId := Map()
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

ShortTitle(title, maxLen := 28) {
    if StrLen(title) <= maxLen
        return title
    return SubStr(title, 1, maxLen - 1) . "..."
}

CycleTabs(direction) {
    global g_TabOrder, g_ActiveTabId

    count := g_TabOrder.Length
    if count < 2
        return

    currentIndex := 0
    for idx, tabId in g_TabOrder {
        if tabId = g_ActiveTabId {
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

    SelectTab(g_TabOrder[nextIndex])
}

StackTabsHostIsActive() {
    global g_HostHwnd

    if !g_HostHwnd
        return false

    return WinActive("ahk_id " g_HostHwnd)
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

TryDeleteMapKey(mapObj, key) {
    if mapObj.Has(key)
        mapObj.Delete(key)
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


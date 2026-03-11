; StackTabs - owner-aware embedded ticket host
; AutoHotkey v2

#Requires AutoHotkey v2.0

; ============ CONFIGURATION ============
; Title text that must appear in the popup/shell window.
g_WindowTitleMatch := "Powershell"

; Optional EXE filter. Leave blank to match any process.
g_TargetExe := ""

; How often to scan for new / replaced windows.
g_RefreshInterval := 500
g_CaptureDelayMs := 900
g_TabDisappearGraceMs := 300

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
g_TabSlotMax := 20
g_CloseButtonWidth := 22
g_PopoutButtonWidth := 22

; Diagnostics.
g_DebugLogPath := A_ScriptDir "\StackTabs-discovery.txt"

; ============ STATE ============
g_MainHost := ""             ; HostInstance for main window
g_PopoutHosts := []          ; array of HostInstance for popped-out windows
g_PendingCandidates := Map() ; tabId -> {firstSeen, candidate} (main host only)
g_IsCleaningUp := false

BuildHostInstance(false)  ; create main host
OnExit(CleanupAll)
RefreshWindows()
SetTimer(RefreshWindows, g_RefreshInterval)

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

; Win+Shift+R restores all windows and exits (use before reload to avoid work loss).
#+r:: {
    CleanupAll()
    ExitApp()
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

IsWindowExists(hwnd) {
    return !!DllCall("IsWindow", "ptr", hwnd, "int")
}

GetClientWidth(hwnd) {
    try {
        clientX := 0
        clientY := 0
        clientW := 0
        clientH := 0
        WinGetClientPos(&clientX, &clientY, &clientW, &clientH, "ahk_id " hwnd)
        return clientW
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
    for host in GetAllHosts() {
        if host.hwnd = hwnd
            return host
    }
    return ""
}

BuildHostInstance(isPopout := false) {
    global g_MainHost, g_PopoutHosts
    global g_HostTitle, g_HostWidth, g_HostHeight, g_HostMinWidth, g_HostMinHeight
    global g_TabHeight, g_CloseButtonWidth

    host := Object()
    host.isPopout := isPopout
    host.tabRecords := Map()
    host.tabOrder := []
    host.hwndToTabId := Map()
    host.activeTabId := ""
    host.tabButtons := Map()
    host.tabCloseButtons := Map()
    host.tabSlotButtons := []
    host.tabSlotCloseButtons := []
    host.tabSlotPopoutButtons := []

    title := isPopout ? (g_HostTitle " (popped out)") : g_HostTitle
    host.gui := Gui("+Resize +MinSize" g_HostMinWidth "x" g_HostMinHeight, title)
    host.gui.BackColor := "1E1E1E"
    host.gui.MarginX := 0
    host.gui.MarginY := 0
    host.gui.SetFont("s10 cWhite", "Segoe UI")
    host.gui.OnEvent("Close", HostGuiClosed.Bind(host))
    host.gui.OnEvent("Size", HostGuiResized.Bind(host))

    host.statusText := host.gui.Add("Text", "Hidden x0 y0 w0 h0", "")
    host.hwnd := host.gui.Hwnd
    host.clientHwnd := host.hwnd
    host.gui.Show("w" g_HostWidth " h" g_HostHeight)

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
        host.gui.Destroy()
    } else {
        CleanupAll()
        ExitApp()
    }
}

HostGuiResized(host, guiObj, minMax, width, height) {
    if minMax = -1
        return
    LayoutTabButtons(host, width)
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
                    if CreateTrackedTab(host, candidate)
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
        if embeddedHwnds.Has(candidate.contentHwnd "")
            continue
        if embeddedHwnds.Has(candidate.topHwnd "")
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

CreateTrackedTab(host, candidate) {
    if host.tabRecords.Has(candidate.id)
        return false

    record := BuildTrackedRecord(candidate)
    host.tabRecords[candidate.id] := record

    if !AttachTrackedWindow(host, candidate.id) {
        host.tabRecords.Delete(candidate.id)
        return false
    }

    IndexTrackedHwnds(host, candidate.id)
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
    UnindexTrackedHwnds(host, tabId)

    record := BuildTrackedRecord(candidate)
    host.tabRecords[tabId] := record
    AttachTrackedWindow(host, tabId)
    IndexTrackedHwnds(host, tabId)
    AppendDebugLog("Rebound tab: " tabId "`r`n" candidate.hierarchySummary "`r`n")
}

CloseTab(host, tabId) {
    if !host.tabRecords.Has(tabId)
        return

    record := host.tabRecords[tabId]
    topHwnd := record.topHwnd
    RemoveTrackedTab(host, tabId, false)
    if WinExist("ahk_id " topHwnd)
        WinClose("ahk_id " topHwnd)

    SetTimer(CloseTabDeferredUpdate.Bind(host), -1)
}

CloseTabDeferredUpdate(host, *) {
    LayoutTabButtons(host)
    ShowOnlyActiveTab(host)
    UpdateHostTitle(host)
    RedrawHostWindow(host)
}

CloseActiveTab(host) {
    if host.activeTabId != ""
        CloseTab(host, host.activeTabId)
}

RemoveTrackedTab(host, tabId, restoreWindow := true) {
    if !host.tabRecords.Has(tabId)
        return

    DetachTrackedWindow(host, tabId, restoreWindow, true)
    UnindexTrackedHwnds(host, tabId)

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

IndexTrackedHwnds(host, tabId) {
    if !host.tabRecords.Has(tabId)
        return

    record := host.tabRecords[tabId]
    host.hwndToTabId[record.contentHwnd] := tabId
    if record.topHwnd != record.contentHwnd
        host.hwndToTabId[record.topHwnd] := tabId
}

UnindexTrackedHwnds(host, tabId) {
    if !host.tabRecords.Has(tabId)
        return

    record := host.tabRecords[tabId]
    TryDeleteMapKey(host.hwndToTabId, record.contentHwnd)
    if record.topHwnd != record.contentHwnd
        TryDeleteMapKey(host.hwndToTabId, record.topHwnd)
}

LayoutTabButtons(host, windowWidth := 0) {
    global g_HostWidth, g_HostPadding, g_TabGap, g_MinTabWidth, g_MaxTabWidth, g_TabHeight
    global g_CloseButtonWidth, g_PopoutButtonWidth, g_TabSlotMax

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

    needed := Min(tabCount, g_TabSlotMax)
    while host.tabSlotButtons.Length < needed {
        btn := host.gui.Add("Text", "Hidden x0 y7 w100 h" g_TabHeight " +0x200 +0x100 Border Center", "")
        btn.SetFont("s9", "Segoe UI Semibold")
        host.tabSlotButtons.Push(btn)
        popoutBtn := host.gui.Add("Text", "Hidden x0 y7 w" g_PopoutButtonWidth " h" g_TabHeight " +0x200 +0x100 Border Center", "")
        popoutBtn.SetFont("s9", "Segoe UI")
        popoutBtn.Opt("cAAAAAA")
        host.tabSlotPopoutButtons.Push(popoutBtn)
        closeBtn := host.gui.Add("Text", "Hidden x0 y7 w" g_CloseButtonWidth " h" g_TabHeight " +0x200 +0x100 Border Center", "×")
        closeBtn.SetFont("s10 Bold", "Segoe UI")
        closeBtn.Opt("cAAAAAA")
        host.tabSlotCloseButtons.Push(closeBtn)
    }

    usableWidth := Max(200, windowWidth - (g_HostPadding * 2))
    tabWidth := Floor((usableWidth - ((tabCount - 1) * g_TabGap)) / tabCount)
    tabWidth := Max(g_MinTabWidth, Min(g_MaxTabWidth, tabWidth))
    titleWidth := tabWidth - extraBtnWidth

    host.tabButtons := Map()
    host.tabCloseButtons := Map()
    host.tabPopoutButtons := Map()
    x := g_HostPadding
    for i, tabId in host.tabOrder {
        if i > g_TabSlotMax
            break
        btn := host.tabSlotButtons[i]
        popoutBtn := host.tabSlotPopoutButtons[i]
        closeBtn := host.tabSlotCloseButtons[i]
        title := host.tabRecords.Has(tabId) ? host.tabRecords[tabId].title : "Window"
        btn.Text := ShortTitle(title, 20)
        btn.Move(x, 7, titleWidth, g_TabHeight)
        btn.OnEvent("Click", SelectTab.Bind(host, tabId))
        btn.Visible := true
        host.tabButtons[tabId] := btn

        popoutBtn.Move(x + titleWidth, 7, g_PopoutButtonWidth, g_TabHeight)
        if host.isPopout {
            popoutBtn.Text := "←"
            popoutBtn.OnEvent("Click", MergeBackClick.Bind(host, tabId))
        } else {
            popoutBtn.Text := "↗"
            popoutBtn.OnEvent("Click", PopOutClick.Bind(host, tabId))
        }
        popoutBtn.Visible := true
        host.tabPopoutButtons[tabId] := popoutBtn

        closeBtn.Move(x + titleWidth + g_PopoutButtonWidth, 7, g_CloseButtonWidth, g_TabHeight)
        closeBtn.OnEvent("Click", CloseTabClick.Bind(host, tabId))
        closeBtn.Visible := true
        host.tabCloseButtons[tabId] := closeBtn

        x += tabWidth + g_TabGap
    }

    Loop Max(0, host.tabSlotButtons.Length - tabCount) {
        i := tabCount + A_Index
        host.tabSlotButtons[i].Visible := false
        host.tabSlotPopoutButtons[i].Visible := false
        host.tabSlotCloseButtons[i].Visible := false
    }

    UpdateTabButtonStyles(host)
}

CloseTabClick(host, tabId, *) {
    CloseTab(host, tabId)
}

PopOutClick(host, tabId, *) {
    PopOutTab(host, tabId)
}

MergeBackClick(host, tabId, *) {
    MergeBackTab(host, tabId)
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
    ; Detach from source (don't restore - we're moving to new host)
    DetachTrackedWindow(sourceHost, tabId, false, false)
    UnindexTrackedHwnds(sourceHost, tabId)
    for idx, currentId in sourceHost.tabOrder {
        if currentId = tabId {
            sourceHost.tabOrder.RemoveAt(idx)
            break
        }
    }
    sourceHost.tabRecords.Delete(tabId)
    if sourceHost.activeTabId = tabId
        sourceHost.activeTabId := sourceHost.tabOrder.Length ? sourceHost.tabOrder[1] : ""

    ; Create pop-out host and attach the tab
    popoutHost := BuildHostInstance(true)
    popoutHost.tabRecords[tabId] := record
    popoutHost.tabOrder.Push(tabId)
    popoutHost.activeTabId := tabId

    if !AttachTrackedWindow(popoutHost, tabId) {
        ; Failed - restore to source
        sourceHost.tabRecords[tabId] := record
        sourceHost.tabOrder.Push(tabId)
        sourceHost.activeTabId := tabId
        AttachTrackedWindow(sourceHost, tabId)
        IndexTrackedHwnds(sourceHost, tabId)
        g_PopoutHosts.Pop()
        popoutHost.gui.Destroy()
        return
    }
    IndexTrackedHwnds(popoutHost, tabId)

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
    DetachTrackedWindow(popoutHost, tabId, false, false)
    UnindexTrackedHwnds(popoutHost, tabId)
    for idx, currentId in popoutHost.tabOrder {
        if currentId = tabId {
            popoutHost.tabOrder.RemoveAt(idx)
            break
        }
    }
    popoutHost.tabRecords.Delete(tabId)

    ; Add to main host
    g_MainHost.tabRecords[tabId] := record
    g_MainHost.tabOrder.Push(tabId)
    if (g_MainHost.activeTabId = "")
        g_MainHost.activeTabId := tabId

    if !AttachTrackedWindow(g_MainHost, tabId) {
        ; Failed - put back in popout
        popoutHost.tabRecords[tabId] := record
        popoutHost.tabOrder.Push(tabId)
        popoutHost.activeTabId := tabId
        AttachTrackedWindow(popoutHost, tabId)
        IndexTrackedHwnds(popoutHost, tabId)
        return
    }
    IndexTrackedHwnds(g_MainHost, tabId)

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
        ; Place host2 to the right of host1 with small gap
        gap := 8
        x2 := x1 + w1 + gap
        host2.gui.Show("x" x2 " y" y1 " w" w1 " h" h1)
    } catch {
        ; Fallback: just show at default position
        host2.gui.Show()
    }
}

ShowOnlyActiveTab(host) {
    if (host.activeTabId != "") && !host.tabRecords.Has(host.activeTabId)
        host.activeTabId := ""
    if (host.activeTabId = "") && host.tabOrder.Length
        host.activeTabId := host.tabOrder[1]

    if host.activeTabId = "" {
        UpdateTabButtonStyles(host)
        return
    }

    GetEmbedRect(host, &areaX, &areaY, &areaW, &areaH)

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
    for tabId, ctrl in host.tabButtons {
        title := host.tabRecords.Has(tabId) ? ShortTitle(host.tabRecords[tabId].title, 20) : "Window"
        if tabId = host.activeTabId {
            ctrl.Text := title
            ctrl.SetFont("s9 Bold", "Segoe UI Semibold")
            ctrl.Opt("Background0x2D7DFF cFFFFFF")
            if host.tabCloseButtons.Has(tabId)
                host.tabCloseButtons[tabId].Opt("Background0x2D7DFF cFFFFFF")
            if host.tabPopoutButtons.Has(tabId)
                host.tabPopoutButtons[tabId].Opt("Background0x2D7DFF cFFFFFF")
        } else {
            ctrl.Text := title
            ctrl.SetFont("s9 Norm", "Segoe UI")
            ctrl.Opt("Background0x30343B cD8DEE9")
            if host.tabCloseButtons.Has(tabId)
                host.tabCloseButtons[tabId].Opt("Background0x30343B cAAAAAA")
            if host.tabPopoutButtons.Has(tabId)
                host.tabPopoutButtons[tabId].Opt("Background0x30343B cAAAAAA")
        }
    }
}

; Placeholder for future status-bar display; g_StatusText control exists but is hidden.
UpdateStatusText() {
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
        host.gui.Title := g_HostTitle . " (" liveCount ") - " . host.tabRecords[host.activeTabId].title . suffix
    else
        host.gui.Title := g_HostTitle . " (" liveCount ")" . suffix
}

GetEmbedRect(host, &x, &y, &w, &h) {
    global g_HostWidth, g_HostHeight, g_HostPadding, g_HeaderHeight

    x := g_HostPadding
    y := g_HeaderHeight + g_HostPadding

    if host.hwnd && WinExist("ahk_id " host.hwnd) {
        try {
            clientX := 0
            clientY := 0
            clientW := 0
            clientH := 0
            WinGetClientPos(&clientX, &clientY, &clientW, &clientH, "ahk_id " host.hwnd)
            w := Max(200, clientW - (g_HostPadding * 2))
            h := Max(140, clientH - y - g_HostPadding)
            return
        }
    }

    w := Max(200, g_HostWidth - (g_HostPadding * 2))
    h := Max(140, g_HostHeight - y - g_HostPadding)
}

CleanupAll(*) {
    global g_MainHost, g_PopoutHosts, g_IsCleaningUp, g_PendingCandidates

    if g_IsCleaningUp
        return

    g_IsCleaningUp := true

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
    activeHwnd := WinGetID("A")
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

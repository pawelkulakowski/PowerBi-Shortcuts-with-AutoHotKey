; Power BI – Switch Views on the left rail by coordinates (Ctrl+Tab / Ctrl+Shift+Tab)
; AutoHotkey v2
#Requires AutoHotkey v2.0
#SingleInstance Force

; ====== config / storage ======
IniPath := A_ScriptDir "\pbiviews.ini"

; Keep the same “order matters” idea as your original Views_All, but now it's keys (names) instead of image candidates
Views_All := [
    "Report",  ; Report
    "Table",   ; Table
    "Model",   ; Model
    "DAX",     ; DAX
    "TMDL"     ; TMDL
]

; ---- left rail search band (kept so your window math remains familiar; not strictly required for coords) ----
RAIL_LeftPad   := 5
RAIL_Width     := 50
RAIL_TopPad    := 180
RAIL_BottomPad := 600

; click offset is not needed for coords, but we'll keep the variables for minimal diffs
CLICK_OffX := 0
CLICK_OffY := 0

; Debug marker (used ONLY in calibration)
DBG_Size  := 14
DBG_Fade  := 450
DBG_Alpha := 200

; Active-view detection settings
COLOR_TOL := 40                                 ; increase to ~60 if theme varies
SAMPLE_OFFSETS := [[0,0],[2,0],[0,2],[2,2]]     ; average 2x2 pixels to be robust

; ====== scope hotkeys to Power BI only ======
#HotIf WinActive("ahk_exe PBIDesktop.exe")
^Tab::  CycleViews( 1)    ; forward
^+Tab:: CycleViews(-1)    ; backward
#q::    CalibrateAll()    ; Win+Q → calibration wizard
#HotIf

; ====== helpers: INI I/O ======
GetCoord(key) {
    global IniPath
    if !FileExist(IniPath)
        return ""
    x := IniRead(IniPath, "coords", key "X", "")
    y := IniRead(IniPath, "coords", key "Y", "")
    return (x != "" && y != "") ? [x+0, y+0] : ""
}

SetCoord(key, x, y) {
    global IniPath
    IniWrite(x, IniPath, "coords", key "X")
    IniWrite(y, IniPath, "coords", key "Y")
}

GetActiveColor(key) {
    global IniPath
    if !FileExist(IniPath)
        return ""
    c := IniRead(IniPath, "coords", key "C", "")
    return (c != "") ? (c+0) : ""
}

SetActiveColor(key, color) {
    global IniPath
    IniWrite(color, IniPath, "coords", key "C")
}

; ====== helpers: index wrap (unchanged) ======
NextIndex(idx, dir, count) {
    idx0 := (idx - 1) + dir
    idx0 := Mod(idx0, count)
    if (idx0 < 0)
        idx0 += count
    return idx0 + 1
}

; ====== color helpers for active detection ======
ColorToRGB(color) {
    r := (color >> 16) & 0xFF
    g := (color >> 8)  & 0xFF
    b :=  color        & 0xFF
    return [r,g,b]
}
RGBToColor(r,g,b) {
    return (r<<16) | (g<<8) | b
}
ColorDistance(c1, c2) {
    rgb1 := ColorToRGB(c1), rgb2 := ColorToRGB(c2)
    dr := rgb1[1]-rgb2[1], dg := rgb1[2]-rgb2[2], db := rgb1[3]-rgb2[3]
    return Sqrt(dr*dr + dg*dg + db*db)
}
SampleColor(x, y) {
    totalR := 0, totalG := 0, totalB := 0, n := 0
    for off in SAMPLE_OFFSETS {
        cx := x + off[1], cy := y + off[2]
        c := PixelGetColor(cx, cy, "RGB")
        rgb := ColorToRGB(c)
        totalR += rgb[1], totalG += rgb[2], totalB += rgb[3]
        n += 1
    }
    return RGBToColor(Round(totalR/n), Round(totalG/n), Round(totalB/n))
}

; ====== helpers: debug marker ======
DebugMark(x, y, size := 12, ms := 400, alpha := 200) {
    g := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20") ; click-through
    g.BackColor := "Red"
    hwnd := g.Hwnd
    g.Show("x" (x - Floor(size/2)) " y" (y - Floor(size/2)) " w" size " h" size " NoActivate")
    WinSetTransparent(alpha, "ahk_id " hwnd)
    SetTimer(() => g.Destroy(), -ms)
}

; ====== coords clicker (replaces ImageSearch) ======
ClickCoordForKey(key, wx, wy) {
    coords := GetCoord(key)
    if !coords
        return false

    clickX := wx + coords[1] + CLICK_OffX
    clickY := wy + coords[2] + CLICK_OffY

    Click(clickX, clickY)
    ; NO DebugMark here → normal operation stays clean/fast
    return true
}

; ====== detect which view is active (by color at stored coords) ======
DetectActiveIndex(wx, wy) {
    global Views_All, COLOR_TOL
    for i, key in Views_All {
        coords := GetCoord(key)
        stored := GetActiveColor(key)
        if !coords || (stored = "")
            continue
        px := wx + coords[1]
        py := wy + coords[2]
        live := SampleColor(px, py)
        if (ColorDistance(live, stored) <= COLOR_TOL)
            return i
    }
    return 0
}

; ====== core ======
CycleViews(direction := 1) {
    global Views_All, RAIL_LeftPad, RAIL_Width, RAIL_TopPad, RAIL_BottomPad
    static idx := 0  ; will be advanced to 1..N on first call

    if !WinExist("ahk_exe PBIDesktop.exe")
        return

    SendMode "Input"
    SetWinDelay(-1), SetControlDelay(-1), SetKeyDelay(-1,-1)
    SetMouseDelay(-1), SetDefaultMouseSpeed(0)
    CoordMode "Mouse", "Screen"
    CoordMode "Pixel", "Screen"

    MouseGetPos &oldX, &oldY
    try {
        WinActivate "ahk_exe PBIDesktop.exe"
        WinWaitActive "ahk_exe PBIDesktop.exe"
        WinGetPos &x, &y, &w, &h, "ahk_exe PBIDesktop.exe"

        bandLeft   := x + RAIL_LeftPad
        bandRight  := bandLeft + RAIL_Width
        bandTop    := y + RAIL_TopPad
        bandBottom := y + h - RAIL_BottomPad

        count := Views_All.Length
        if (count = 0)
            return

        ; NEW: detect the *actual* active view each time (loop-safe)
        activeIdx := DetectActiveIndex(x, y)
        if (activeIdx > 0)
            idx := activeIdx
        else if (idx = 0)
            idx := 1

        ; move once in requested direction (1-based safe)
        idx := NextIndex(idx, direction, count)

        ; Try up to 'count' entries, wrapping, until one is calibrated and clicked
        attempts := 0
        while (attempts < count) {
            key := Views_All[idx]  ; string key (e.g. "Report")
            if ClickCoordForKey(key, x, y) {
                ; success → leave idx where it landed; next call detects again
                return
            }
            ; not calibrated → move again in same direction
            idx := NextIndex(idx, direction, count)
            attempts += 1
        }
        ; silent if nothing matched
        ; TrayTip("Power BI View Switch", "No calibrated coordinates. Press Win+Q to calibrate.", 2000)
    } finally {
        MouseMove oldX, oldY, 0
    }
}

; ====== calibration wizard (Win+Q) ======
CalibrateAll() {
    global Views_All
    if !WinExist("ahk_exe PBIDesktop.exe") {
        MsgBox("Power BI Desktop window not found. Open it first.", "Calibration", "Icon!")
        return
    }

    WinActivate("ahk_exe PBIDesktop.exe")
    WinWaitActive("ahk_exe PBIDesktop.exe")

    CoordMode("Mouse", "Screen")
    CoordMode("Pixel", "Screen")
    WinGetPos(&wx, &wy, &ww, &wh, "ahk_exe PBIDesktop.exe")

    MsgBox("Calibration will record, for each view, its icon position and ACTIVE color."
        . "`n`nFor each view:"
        . "`n  1) Hover the mouse over the icon → press ENTER to capture position (ESC to skip)."
        . "`n  2) Click the icon to activate that view → press ENTER to capture its active color."
        . "`n`nTip: You can re-run Win+Q anytime to update.", "Calibration", "Iconi")

    for key in Views_All {
        ; Step 1: position
        resp := MsgBox("Hover over the '" key "' icon, then press ENTER to capture position."
            . "`n(ESC to skip this view.)", "Calibrate position: " key, "Iconi 1 1")

        if (resp = "OK") {
            MouseGetPos(&mx, &my)
            offX := mx - wx
            offY := my - wy
            SetCoord(key, offX, offY)
            DebugMark(mx, my, DBG_Size*1.3, DBG_Fade+200, DBG_Alpha)
            TrayTip("Saved " key, "X=" offX "  Y=" offY, 1000)

            ; Step 2: active color
            resp2 := MsgBox("Now CLICK the '" key "' icon to make it ACTIVE, then press ENTER to capture color."
                . "`n(ESC to skip color capture.)", "Capture active color: " key, "Iconi 1 1")
            if (resp2 = "OK") {
                px := wx + offX, py := wy + offY
                color := SampleColor(px, py)
                SetActiveColor(key, color)
                TrayTip("Saved active color for " key, Format("0x{:06X}", color), 1200)
            } else {
                TrayTip("Saved position only for " key, "", 800)
            }
        } else {
            TrayTip("Skipped " key, "", 800)
        }
        Sleep 150
    }

    MsgBox("Calibration complete. Use Ctrl+Tab / Ctrl+Shift+Tab to switch views."
        . "`nRed squares appear only during calibration; normal operation is silent.", "Calibration", "Iconi")
}

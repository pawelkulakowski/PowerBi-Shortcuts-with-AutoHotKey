; Power BI – Switch Views on the left rail by image (Ctrl+Tab / Ctrl+Shift+Tab)
; AutoHotkey v2
#Requires AutoHotkey v2.0
#SingleInstance Force

ImgDir := A_ScriptDir "\img\"

; ---- images (order matters) ----
Views_All := [
    [ImgDir "view_report.png"],   ; Report
    [ImgDir "view_table.png"],    ; Table
    [ImgDir "view_model.png"],    ; Model
    [ImgDir "view_dax.png"],      ; DAX
    [ImgDir "view_tmdl.png"]      ; TMDL
]

; Active-state images (same order)
Views_Active := [
    [ImgDir "view_report_active.png"],
    [ImgDir "view_table_active.png"],
    [ImgDir "view_model_active.png"],
    [ImgDir "view_dax_active.png"],
    [ImgDir "view_tmdl_active.png"]
]

; ---- scope hotkeys to Power BI only ----
#HotIf WinActive("ahk_exe PBIDesktop.exe")
^Tab::  CycleViews( 1)    ; forward
^+Tab:: CycleViews(-1)    ; backward
#HotIf

; ---- left rail search band ----
RAIL_LeftPad   := 5
RAIL_Width     := 50
RAIL_TopPad    := 180
RAIL_BottomPad := 600

CLICK_OffX := 8
CLICK_OffY := 8
SEARCH_TOLS := [90,120,180]

; ---------- helpers ----------
ClickImageInBand(imgPath, x1,y1,x2,y2, tolerances, offX := 8, offY := 8) {
    if !FileExist(imgPath)
        return false
    for tol in tolerances {
        if ImageSearch(&fx, &fy, x1, y1, x2, y2, "*" tol " " imgPath) {
            Click fx + offX, fy + offY
            return true
        }
    }
    return false
}

FindImageInBand(imgPath, x1,y1,x2,y2, tolerances) {
    if !FileExist(imgPath)
        return false
    for tol in tolerances {
        if ImageSearch(&fx, &fy, x1, y1, x2, y2, "*" tol " " imgPath)
            return true
    }
    return false
}

ResolveFirstExisting(paths) {
    for p in paths {
        if FileExist(p)
            return p
    }
    return ""  ; none exist
}

; 1-based index wrap helper → always returns 1..count
NextIndex(idx, dir, count) {
    idx0 := (idx - 1) + dir
    idx0 := Mod(idx0, count)
    if (idx0 < 0)
        idx0 += count
    return idx0 + 1
}

; return 1-based index of active icon, or 0 if none found
DetectActiveIndex(x1,y1,x2,y2) {
    global Views_Active, SEARCH_TOLS
    for idx, variants in Views_Active {
        for img in variants {
            if (img != "" && FindImageInBand(img, x1,y1,x2,y2, SEARCH_TOLS))
                return idx
        }
    }
    return 0
}

; ---------- core ----------
CycleViews(direction := 1) {
    global Views_All, RAIL_LeftPad, RAIL_Width, RAIL_TopPad, RAIL_BottomPad
    global CLICK_OffX, CLICK_OffY, SEARCH_TOLS

    if !WinExist("ahk_exe PBIDesktop.exe")
        return

    SendMode "Input"
    SetWinDelay(-1), SetControlDelay(-1), SetKeyDelay(-1,-1)
    SetMouseDelay(-1), SetDefaultMouseSpeed(0)
    CoordMode "Mouse", "Screen"
    CoordMode "Pixel", "Screen"

    MouseGetPos &oldX, &oldY
    static idx := 0  ; remembered across calls

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

        ; start from currently active icon if detectable
        activeIdx := DetectActiveIndex(bandLeft, bandTop, bandRight, bandBottom)
        if (activeIdx != 0)
            idx := activeIdx

        ; move once in requested direction (1-based safe)
        idx := NextIndex(idx = 0 ? 1 : idx, direction, count)

        ; Try up to 'count' entries, wrapping, until one is found
        attempts := 0
        while (attempts < count) {
            candidates := Views_All[idx]
            candidate := ResolveFirstExisting(candidates)
            if (candidate != "") {
                if ClickImageInBand(candidate, bandLeft, bandTop, bandRight, bandBottom,
                                    SEARCH_TOLS, CLICK_OffX, CLICK_OffY) {
                    return  ; success
                }
            }
            idx := NextIndex(idx, direction, count)
            attempts += 1
        }
        ; silent if nothing matched
    } finally {
        MouseMove oldX, oldY, 0
    }
}

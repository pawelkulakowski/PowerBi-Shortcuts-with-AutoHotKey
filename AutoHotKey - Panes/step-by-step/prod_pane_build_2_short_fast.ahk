; Power BI – Build/Format pane TOGGLE (Win+2) + TITLE CLICK (Win+W) — ultra-fast
; AutoHotkey v2
#Requires AutoHotkey v2.0
#SingleInstance Force

; ---------- images ----------
ImgPaneIcon := A_ScriptDir "\img\pane_build.png"    ; right-edge icon
ImgTitleRow := A_ScriptDir "\img\title_build.png"   ; "Build/Format" title row

; ---------- hotkeys ----------
#2:: TogglePane_ByImages()
#w:: Title_FindAndClick()

; ---------- search bands (tweak if needed) ----------
ICON_RightMargin := 10
ICON_BandWidth   := 120
ICON_TopOffset   := 100
ICON_BottomPad   := 600

TITLE_RightMargin := 50
TITLE_BandWidth   := 1400
TITLE_TopOffset   := 150
TITLE_BandHeight  := 80

; Click offset
CLICK_OffX := 10
CLICK_OffY := 10

; Fast tolerances
SEARCH_TOLS := [40,70,90]

; ---------- helpers (minimal) ----------
ClickImageInBand(imgPath, x1,y1,x2,y2, tolerances, offX := 10, offY := 10) {
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

FindImageInBand(imgPath, x1,y1,x2,y2, tolerances, &outX?, &outY?) {
    if !FileExist(imgPath)
        return false
    for tol in tolerances {
        if ImageSearch(&fx, &fy, x1, y1, x2, y2, "*" tol " " imgPath) {
            outX := fx, outY := fy
            return true
        }
    }
    return false
}

; ---------- MAIN TOGGLE (Win+2) ----------
TogglePane_ByImages() {
    global ImgPaneIcon, ImgTitleRow
    global ICON_RightMargin, ICON_BandWidth, ICON_TopOffset, ICON_BottomPad
    global TITLE_RightMargin, TITLE_BandWidth, TITLE_TopOffset, TITLE_BandHeight
    global CLICK_OffX, CLICK_OffY, SEARCH_TOLS

    static running := false
    if running
        return
    running := true
    Critical "On"

    if !WinExist("ahk_exe PBIDesktop.exe") {
        running := false
        Critical "Off"
        return
    }

    ; speed + absolute coords
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

        ; bands
        rIcon := x + w - ICON_RightMargin
        lIcon := rIcon - ICON_BandWidth
        tIcon := y + ICON_TopOffset
        bIcon := y + h - ICON_BottomPad

        rTitle := x + w - TITLE_RightMargin
        lTitle := rTitle - TITLE_BandWidth
        tTitle := y + TITLE_TopOffset
        bTitle := tTitle + TITLE_BandHeight

        ; TITLE FIRST
        titleVisible := FindImageInBand(ImgTitleRow, lTitle,tTitle,rTitle,bTitle, SEARCH_TOLS, &tx, &ty)

        if titleVisible {
            ; Pane open -> close via icon
            ClickImageInBand(ImgPaneIcon, lIcon,tIcon,rIcon,bIcon, SEARCH_TOLS, CLICK_OffX, CLICK_OffY)
            return
        } else {
            ; Pane closed -> open via icon, then hit title
            if ClickImageInBand(ImgPaneIcon, lIcon,tIcon,rIcon,bIcon, SEARCH_TOLS, CLICK_OffX, CLICK_OffY) {
                ; quick retries without sleeps
                Loop 8 {
                    if ClickImageInBand(ImgTitleRow, lTitle,tTitle,rTitle,bTitle, SEARCH_TOLS, CLICK_OffX, CLICK_OffY)
                        break
                }
            }
            return
        }
    } finally {
        MouseMove oldX, oldY, 0
        Critical "Off"
        running := false
    }
}

; ---------- TITLE CLICK ONLY (Win+W) ----------
Title_FindAndClick() {
    global ImgTitleRow
    global TITLE_RightMargin, TITLE_BandWidth, TITLE_TopOffset, TITLE_BandHeight
    global CLICK_OffX, CLICK_OffY, SEARCH_TOLS

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

        rTitle := x + w - TITLE_RightMargin
        lTitle := rTitle - TITLE_BandWidth
        tTitle := y + TITLE_TopOffset
        bTitle := tTitle + TITLE_BandHeight

        ClickImageInBand(ImgTitleRow, lTitle,tTitle,rTitle,bTitle, SEARCH_TOLS, CLICK_OffX, CLICK_OffY)
    } finally {
        MouseMove oldX, oldY, 0
    }
}

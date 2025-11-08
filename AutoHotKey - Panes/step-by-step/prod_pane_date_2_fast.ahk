; Power BI – Data pane TOGGLE (Win+1): ultra-fast image-driven
; AutoHotkey v2
#Requires AutoHotkey v2.0
#SingleInstance Force

; ---------- images ----------
ImgPaneIcon := A_ScriptDir "\img\pane_data.png"    ; right-edge icon for Data pane rail
ImgTitleRow := A_ScriptDir "\img\title_data.png"   ; "Data" title row crop

; ---------- hotkey ----------
#1:: TogglePane_ByImages()

; ---------- search bands (tweak as needed) ----------
; Right-edge icon band
ICON_RightMargin := 10
ICON_BandWidth   := 120
ICON_TopOffset   := 100
ICON_BottomPad   := 600

; Title band (only the title row area)
TITLE_RightMargin := 50
TITLE_BandWidth   := 1400
TITLE_TopOffset   := 150
TITLE_BandHeight  := 80

; Click offset inside match
CLICK_OffX := 10
CLICK_OffY := 10

; Fewer, faster tolerances
SEARCH_TOLS := [40, 70, 90]

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

; ---------- core ----------
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

    ; remember mouse
    MouseGetPos &oldX, &oldY

    try {
        ; activate + get window rect
        WinActivate "ahk_exe PBIDesktop.exe"
        WinWaitActive "ahk_exe PBIDesktop.exe"
        WinGetPos &x, &y, &w, &h, "ahk_exe PBIDesktop.exe"

        ; compute bands
        rIcon := x + w - ICON_RightMargin
        lIcon := rIcon - ICON_BandWidth
        tIcon := y + ICON_TopOffset
        bIcon := y + h - ICON_BottomPad

        rTitle := x + w - TITLE_RightMargin
        lTitle := rTitle - TITLE_BandWidth
        tTitle := y + TITLE_TopOffset
        bTitle := tTitle + TITLE_BandHeight

        ; --- TITLE FIRST (pane open?) ---
        titleVisible := FindImageInBand(ImgTitleRow, lTitle,tTitle,rTitle,bTitle, SEARCH_TOLS, &tx, &ty)

        if titleVisible {
            ; Pane is OPEN -> click ICON to CLOSE
            ClickImageInBand(ImgPaneIcon, lIcon,tIcon,rIcon,bIcon, SEARCH_TOLS, CLICK_OffX, CLICK_OffY)
            return
        } else {
            ; Pane is CLOSED -> click ICON to OPEN, then try TITLE quickly (no sleeps)
            if ClickImageInBand(ImgPaneIcon, lIcon,tIcon,rIcon,bIcon, SEARCH_TOLS, CLICK_OffX, CLICK_OffY) {
                ; brief busy-wait retries for title click
                Loop 8 {
                    if ClickImageInBand(ImgTitleRow, lTitle,tTitle,rTitle,bTitle, SEARCH_TOLS, CLICK_OffX, CLICK_OffY)
                        break
                }
            }
            return
        }
    } finally {
        ; restore mouse & clear guard
        MouseMove oldX, oldY, 0
        Critical "Off"
        running := false
    }
}

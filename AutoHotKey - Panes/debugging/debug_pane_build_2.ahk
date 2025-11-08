; Power BI – Build/Format pane TOGGLE (Win+1): image-driven
; Flow:
;   1) Search title_build.png FIRST. (Show yellow band over title area while searching.)
;   2) If found -> pane OPEN -> search/click pane_build.png to CLOSE (show icon band while searching), restore mouse.
;   3) If not found -> pane CLOSED -> search/click pane_build.png to OPEN (show icon band), 
;      then search/click title_build.png (show title band), restore mouse.
; Visuals: only yellow band for the CURRENT search; red dot where we click.
; AutoHotkey v2
#Requires AutoHotkey v2.0
#SingleInstance Force

; ---------- images ----------
ImgPaneIcon   := A_ScriptDir "\img\pane_build.png"   ; right-edge icon (Build/Format rail)
ImgTitleRow   := A_ScriptDir "\img\title_build.png"  ; title row crop ("Build"/"Format")

; ---------- hotkey ----------
#2:: TogglePane_ByImages()

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

; Click offset inside found match (from top-left of image)
CLICK_OffX := 10
CLICK_OffY := 10

; Tolerances to try for ImageSearch (light -> heavy)
SEARCH_TOLS := [35,55,75,90]

; ---------- helpers ----------
ShowClickDot(x, y, ms := 250, size := 16) {
    g := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20") ; click-through
    g.BackColor := "Red"
    g.Opt("+LastFound")
    WinSetTransparent(220)
    g.Show("x" (x - size//2) " y" (y - size//2) " w" size " h" size)
    Sleep ms
    try g.Destroy()
}

ShowSearchBand(x1, y1, x2, y2, ms := 500) {
    g := Gui("+AlwaysOnTop -Caption +ToolWindow")
    g.BackColor := "Yellow", g.Opt("+LastFound"), WinSetTransparent(80)
    g.Show("x" x1 " y" y1 " w" (x2 - x1) " h" (y2 - y1))
    Sleep ms
    try g.Destroy()
}

ClickImageInBand(imgPath, x1,y1,x2,y2, tolerances, offX := 10, offY := 10) {
    if !FileExist(imgPath)
        return false
    for tol in tolerances {
        ok := ImageSearch(&fx, &fy, x1, y1, x2, y2, "*" tol " " imgPath)
        if ok {
            cx := fx + offX, cy := fy + offY
            ShowClickDot(cx, cy)
            MouseMove cx, cy, 0
            Click "Left"
            return true
        }
    }
    return false
}

FindImageInBand(imgPath, x1,y1,x2,y2, tolerances, &outX?, &outY?) {
    if !FileExist(imgPath)
        return false
    for tol in tolerances {
        ok := ImageSearch(&fx, &fy, x1, y1, x2, y2, "*" tol " " imgPath)
        if ok {
            outX := fx, outY := fy
            return true
        }
    }
    return false
}

TogglePane_ByImages() {
    ; ---- explicit globals ----
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

        ; --- ALWAYS CHECK TITLE FIRST (show only the TITLE band while searching) ---
        ShowSearchBand(lTitle, tTitle, rTitle, bTitle, 400)
        titleVisible := FindImageInBand(ImgTitleRow, lTitle,tTitle,rTitle,bTitle, SEARCH_TOLS, &tx, &ty)

        if titleVisible {
            ; Pane is OPEN -> now search/click ICON to CLOSE
            ShowSearchBand(lIcon, tIcon, rIcon, bIcon, 400)
            ClickImageInBand(ImgPaneIcon, lIcon,tIcon,rIcon,bIcon, SEARCH_TOLS, CLICK_OffX, CLICK_OffY)
            return
        } else {
            ; Pane is CLOSED -> search/click ICON to OPEN
            ShowSearchBand(lIcon, tIcon, rIcon, bIcon, 400)
            if ClickImageInBand(ImgPaneIcon, lIcon,tIcon,rIcon,bIcon, SEARCH_TOLS, CLICK_OffX, CLICK_OffY) {
                ; wait for the UI, then search/click TITLE
                Sleep 150
                ShowSearchBand(lTitle, tTitle, rTitle, bTitle, 400)
                Loop 10 {
                    if ClickImageInBand(ImgTitleRow, lTitle,tTitle,rTitle,bTitle, SEARCH_TOLS, CLICK_OffX, CLICK_OffY)
                        break
                    Sleep 60
                }
            }
            return
        }
    } finally {
        ; ALWAYS restore mouse (even on early return or error)
        MouseMove oldX, oldY, 0
        Critical "Off"
        running := false
    }
}

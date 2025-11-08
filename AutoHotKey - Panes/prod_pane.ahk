; Power BI – Combined Pane Toggles (Win+1/2/3) + Title Clicks (Win+Q/W/E): ultra-fast
; AutoHotkey v2
#Requires AutoHotkey v2.0
#SingleInstance Force

; ---------- search bands ----------
ICON_RightMargin := 10
ICON_BandWidth   := 120
ICON_TopOffset   := 100
ICON_BottomPad   := 600

TITLE_RightMargin := 50
TITLE_BandWidth   := 1400
TITLE_TopOffset   := 150
TITLE_BandHeight  := 80

CLICK_OffX := 10
CLICK_OffY := 10

; Progressive tolerances: start tight for speed
SEARCH_TOLS := [30, 60, 90]

; ---------- hotkeys ----------
; Toggle panes (opens/closes)
#1:: TogglePane("data")
#2:: TogglePane("build")
#3:: TogglePane("format")

; Direct title clicks (just clicks the title)
#q:: ClickTitle("data")
#w:: ClickTitle("build")
#e:: ClickTitle("format")

; ---------- optimized helpers ----------
ClickImageInBand(imgPath, x1, y1, x2, y2, tolerances, offX, offY) {
    Loop tolerances.Length {
        if ImageSearch(&fx, &fy, x1, y1, x2, y2, "*" tolerances[A_Index] " " imgPath) {
            Click fx + offX, fy + offY
            return true
        }
    }
    return false
}

FindImageInBand(imgPath, x1, y1, x2, y2, tolerances) {
    Loop tolerances.Length {
        if ImageSearch(&fx, &fy, x1, y1, x2, y2, "*" tolerances[A_Index] " " imgPath)
            return true
    }
    return false
}

TogglePane(paneType) {
    global ICON_RightMargin, ICON_BandWidth, ICON_TopOffset, ICON_BottomPad
    global TITLE_RightMargin, TITLE_BandWidth, TITLE_TopOffset, TITLE_BandHeight
    global CLICK_OffX, CLICK_OffY, SEARCH_TOLS

    static running := false, imgCache := Map()
    
    if running
        return
    running := true

    ; Pre-compute image paths once
    if !imgCache.Has(paneType) {
        imgCache[paneType] := {
            icon: A_ScriptDir "\img\pane_" paneType ".png",
            title: A_ScriptDir "\img\title_" paneType ".png"
        }
    }
    imgs := imgCache[paneType]

    if !WinExist("ahk_exe PBIDesktop.exe") {
        running := false
        return
    }

    Critical "On"
    SendMode "Input"
    SetWinDelay -1
    SetControlDelay -1
    SetKeyDelay -1, -1
    SetMouseDelay -1
    SetDefaultMouseSpeed 0
    CoordMode "Mouse", "Screen"
    CoordMode "Pixel", "Screen"

    MouseGetPos &oldX, &oldY
    
    WinActivate "ahk_exe PBIDesktop.exe"
    WinGetPos &x, &y, &w, &h, "A"
    
    ; Pre-calculate all bounds
    rIcon := x + w - ICON_RightMargin
    lIcon := rIcon - ICON_BandWidth
    tIcon := y + ICON_TopOffset
    bIcon := y + h - ICON_BottomPad
    
    rTitle := x + w - TITLE_RightMargin
    lTitle := rTitle - TITLE_BandWidth
    tTitle := y + TITLE_TopOffset
    bTitle := tTitle + TITLE_BandHeight

    ; Check if pane is open (title visible)
    if FindImageInBand(imgs.title, lTitle, tTitle, rTitle, bTitle, SEARCH_TOLS) {
        ; Close pane
        ClickImageInBand(imgs.icon, lIcon, tIcon, rIcon, bIcon, SEARCH_TOLS, CLICK_OffX, CLICK_OffY)
    } else {
        ; Open pane and click title
        if ClickImageInBand(imgs.icon, lIcon, tIcon, rIcon, bIcon, SEARCH_TOLS, CLICK_OffX, CLICK_OffY) {
            ; Tight retry loop
            Loop 6 {
                if ClickImageInBand(imgs.title, lTitle, tTitle, rTitle, bTitle, SEARCH_TOLS, CLICK_OffX, CLICK_OffY)
                    break
            }
        }
    }
    
    MouseMove oldX, oldY, 0
    Critical "Off"
    running := false
}

ClickTitle(paneType) {
    global TITLE_RightMargin, TITLE_BandWidth, TITLE_TopOffset, TITLE_BandHeight
    global CLICK_OffX, CLICK_OffY, SEARCH_TOLS

    static imgCache := Map()

    ; Pre-compute image path once
    if !imgCache.Has(paneType) {
        imgCache[paneType] := A_ScriptDir "\img\title_" paneType ".png"
    }
    imgTitle := imgCache[paneType]

    if !WinExist("ahk_exe PBIDesktop.exe")
        return

    SendMode "Input"
    SetWinDelay -1
    SetControlDelay -1
    SetKeyDelay -1, -1
    SetMouseDelay -1
    SetDefaultMouseSpeed 0
    CoordMode "Mouse", "Screen"
    CoordMode "Pixel", "Screen"

    MouseGetPos &oldX, &oldY
    
    WinActivate "ahk_exe PBIDesktop.exe"
    WinGetPos &x, &y, &w, &h, "A"

    rTitle := x + w - TITLE_RightMargin
    lTitle := rTitle - TITLE_BandWidth
    tTitle := y + TITLE_TopOffset
    bTitle := tTitle + TITLE_BandHeight

    ClickImageInBand(imgTitle, lTitle, tTitle, rTitle, bTitle, SEARCH_TOLS, CLICK_OffX, CLICK_OffY)
    
    MouseMove oldX, oldY, 0
}
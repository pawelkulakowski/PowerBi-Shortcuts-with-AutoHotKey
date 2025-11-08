; Power BI – Format pane (Win+2): click by images
; Step 1: pane_format.png (right rail icon)
; Step 2: title_format.png (title row)  ← as requested
; Yellow debug bands + red click dot + retries
; AutoHotkey v2
#Requires AutoHotkey v2.0
#SingleInstance Force

; ---------- images ----------
ImgFormatIcon  := A_ScriptDir "\img\pane_format.png"    ; right-edge Format icon
ImgFormatTitle := A_ScriptDir "\img\title_format.png"   ; "Format" title row crop

; ---------- hotkey ----------
#3:: ActivateFormat_ByImages()

; ---------- bands (tweak if needed) ----------
; Right-edge icon search band
ICON_RightMargin := 10
ICON_BandWidth   := 120
ICON_TopOffset   := 100
ICON_BottomPad   := 600

; Title search band (where the "Format" title lives)
; Make sure this rectangle only covers the title row area
TITLE_RightMargin := 50
TITLE_BandWidth   := 1400
TITLE_TopOffset   := 150   ; move lower -> increase, higher -> decrease
TITLE_BandHeight  := 80   ; thin to avoid the Search box below

; Click offset inside the found image (top-left of match)
CLICK_OffX := 10
CLICK_OffY := 10

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

ActivateFormat_ByImages() {
    global ImgFormatIcon, ImgFormatTitle
    global ICON_RightMargin, ICON_BandWidth, ICON_TopOffset, ICON_BottomPad
    global TITLE_RightMargin, TITLE_BandWidth, TITLE_TopOffset, TITLE_BandHeight
    global CLICK_OffX, CLICK_OffY

    if !WinExist("ahk_exe PBIDesktop.exe")
        return

    ; speed + coords
    SendMode "Input"
    SetWinDelay(-1), SetControlDelay(-1), SetKeyDelay(-1,-1)
    SetMouseDelay(-1), SetDefaultMouseSpeed(0)
    CoordMode "Mouse", "Screen"
    CoordMode "Pixel", "Screen"

    ; remember mouse, activate PBI, window rect
    MouseGetPos &oldX, &oldY
    WinActivate "ahk_exe PBIDesktop.exe"
    WinWaitActive "ahk_exe PBIDesktop.exe"
    WinGetPos &x, &y, &w, &h, "ahk_exe PBIDesktop.exe"

    ; ---------------- Step 1: right-edge icon band ----------------
    r1 := x + w - ICON_RightMargin
    l1 := r1 - ICON_BandWidth
    t1 := y + ICON_TopOffset
    b1 := y + h - ICON_BottomPad

    ; yellow overlay
    g1 := Gui("+AlwaysOnTop -Caption +ToolWindow")
    g1.BackColor := "Yellow", g1.Opt("+LastFound"), WinSetTransparent(80)
    g1.Show("x" l1 " y" t1 " w" r1 - l1 " h" b1 - t1)
    Sleep 300
    try g1.Destroy()

    ; click the format icon
    ClickImageInBand(ImgFormatIcon, l1,t1,r1,b1, [35,55,75,90], CLICK_OffX, CLICK_OffY)

    ; ---------------- Step 2: title band (title_format.png) ----------------
    r2 := x + w - TITLE_RightMargin
    l2 := r2 - TITLE_BandWidth
    t2 := y + TITLE_TopOffset
    b2 := t2 + TITLE_BandHeight

    ; yellow overlay
    g2 := Gui("+AlwaysOnTop -Caption +ToolWindow")
    g2.BackColor := "Yellow", g2.Opt("+LastFound"), WinSetTransparent(80)
    g2.Show("x" l2 " y" t2 " w" r2 - l2 " h" b2 - t2)
    Sleep 300
    try g2.Destroy()

    ; small settle after step 1
    Sleep 120

    ; retry up to ~1s to find and click the title image
    found := false
    Loop 8 {
        if ClickImageInBand(ImgFormatTitle, l2,t2,r2,b2, [35,55,75,90], CLICK_OffX, CLICK_OffY) {
            found := true
            break
        }
        Sleep 120
    }

    ; restore mouse
    MouseMove oldX, oldY, 0
}

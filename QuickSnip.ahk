; QuickSnip.ahk
; Press Alt+Shift+S to snip and copy to clipboard
!+S::
{
    Send, {LWin Down}{Shift Down}{S}{Shift Up}{LWin Up}  ; Open Snip & Sketch
    Sleep, 500
    ClipWait, 2
    if (ErrorLevel) {
        MsgBox, Snipping failed. Try again!
        Return
    }
    MsgBox, Image copied to clipboard!
}
Return

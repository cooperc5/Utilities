; Quick Snip & Copy to Clipboard (Alt + Shift + S)
!+S::
{
    Send, {LWin Down}{Shift Down}{S}{Shift Up}{LWin Up} ; Open Snipping Tool
    Sleep, 500 ; Wait for it to open
    ClipWait, 2 ; Wait for the clipboard to contain an image
    if (ErrorLevel)
    {
        MsgBox, Snipping failed. Try again!
        Return
    }
    MsgBox, Image copied to clipboard! ; Confirmation message
}
Return

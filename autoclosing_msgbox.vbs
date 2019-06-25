'This is useful if you need to display information, but not have a program which prints hundreds of logs to spam with windows.

Call MessageBoxTimer("text to be printed", "title")

Sub MessageBoxTimer( text, title )
    Dim AckTime
	Dim InfoBox
    Set InfoBox = CreateObject("WScript.Shell")
    'Set the message box to close after 10 seconds
    AckTime = 5
    Select Case InfoBox.Popup(text, _
    AckTime, title, 0)
        Case 1, -1
            Exit Sub
    End Select
End Sub

Option Explicit

' Define the colors
Const FOREGROUND_RED = &H0C
Const FOREGROUND_GREEN = &H0A
Const FOREGROUND_BLUE = &H0B

' Get the console object
Dim oCon
Set oCon = CreateObject("WScript.StdOut")

' Set the foreground color
oCon.SetTextColor FOREGROUND_RED

' Write some text to the console
WScript.Echo "This text is red."

' Reset the foreground color
oCon.SetTextColor FOREGROUND_RED Or FOREGROUND_GREEN Or FOREGROUND_BLUE

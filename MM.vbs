
Set WshShell = CreateObject("WScript.Shell")

' Number of iterations (adjust as needed)
maxIterations = 20

' Infinite loop (or adjust conditions as needed)
For i = 1 To maxIterations
    ' Specify the X and Y coordinates where you want the mouse cursor to move
    mouseX = 500
    mouseY = 500

    ' Run the command to move the mouse cursor
    WshShell.Run "powershell -Command ""[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(" & mouseX & ", " & mouseY & ")""", 0, True

    ' Simulate a right-click using keyboard shortcuts
    WshShell.SendKeys "+{F10}"

    ' Introduce a short delay before the left-click
    WScript.Sleep 60000 ' 1,000 milliseconds = 1 second

    ' Simulate a left-click using keyboard shortcuts
    WshShell.SendKeys "{ENTER}"

    ' Display a message indicating the iteration
    'WScript.Echo "Iteration " & i & ": Mouse moved to X=" & mouseX & " Y=" & mouseY & vbCrLf & "Right-click and left-click simulated after a delay"
Next

Set objShell = CreateObject("WScript.Shell")

' Display a dialog box with "Run" and "Exit" buttons
response = MsgBox("Do you want to run the script?", vbQuestion + vbYesNo, "Script Execution")

' Check the user's response
If response = vbYes Then
    ' Batch script content
    batchScript = "@echo off" & vbCrLf & _
        ":LOOP" & vbCrLf & _
        "taskkill /F /IM SmoothScroll.exe" & vbCrLf & _
        "timeout /t 1 /nobreak >nul" & vbCrLf & _
        "rem Start SmoothScroll.exe" & vbCrLf & _
        "start """" ""C:\Users\Tatsh\AppData\Local\SmoothScroll\app-1.2.4.0\SmoothScroll.exe""" & vbCrLf & _
        "timeout /t 1200 /nobreak >nul" & vbCrLf & _
        "goto LOOP"

    ' Save the batch script content to a temporary file
    tempBatchFile = objShell.ExpandEnvironmentStrings("%TEMP%\TempBatchFile.bat")
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.CreateTextFile(tempBatchFile, True)
    objFile.Write batchScript
    objFile.Close

    ' Run the batch script with the window visible
    objShell.Run "cmd /k " & tempBatchFile, 1, True

    ' Clean up the temporary batch file
    objFSO.DeleteFile(tempBatchFile)
Else
    ' User clicked "Exit"
    WScript.Quit
End If

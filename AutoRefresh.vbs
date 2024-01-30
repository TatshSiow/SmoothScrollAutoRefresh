Set objShell = CreateObject("WScript.Shell")

' Display a dialog box with "Run," "Terminate," and "Exit" buttons
response = MsgBox("Yes: Run AutoRefresh" & vbCrLf & "No: Kill Auto Refresh" & vbCrLf & "Cancel: Exit Service", vbYesNoCancel + vbQuestion + vbDefaultButton1, "SmoothScroll AutoRefresh")

' Choose an action based on the user's response
Select Case response
    Case vbYes ' Run
        ' Batch script content
        batchScript = "@echo off" & vbCrLf & _
            ":LOOP" & vbCrLf & _
            "taskkill /F /IM SmoothScroll.exe" & vbCrLf & _
            "timeout /t 1 /nobreak >nul" & vbCrLf & _
            "rem Start SmoothScroll.exe" & vbCrLf & _
            "start """" ""C:\Users\User\AppData\Local\SmoothScroll\app-1.2.4.0\SmoothScroll.exe""" & vbCrLf & _
            "timeout /t 1200 /nobreak >nul" & vbCrLf & _
            "goto LOOP"

        ' Save the batch script content to a temporary file
        tempBatchFile = objShell.ExpandEnvironmentStrings("%TEMP%\TempBatchFile.bat")
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objFile = objFSO.CreateTextFile(tempBatchFile, True)
        objFile.Write batchScript
        objFile.Close

        ' Run the batch script with the window visible
        objShell.Run "cmd /k " & tempBatchFile, 0, True

        ' Clean up the temporary batch file
        objFSO.DeleteFile(tempBatchFile)

    Case vbNo ' Terminate
        ' Terminate openconsole.exe and cmd.exe
        objShell.Run "taskkill /F /IM openconsole.exe", 0, True
        objShell.Run "taskkill /F /IM cmd.exe", 0, True


    Case vbCancel ' Exit
        ' Do nothing or add any cleanup logic if needed

End Select

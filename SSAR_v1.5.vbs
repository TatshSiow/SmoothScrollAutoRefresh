Set objShell = CreateObject("WScript.Shell")

smoothScrollPath = FindSmoothScrollPath()

response = MsgBox("Yes: Run AutoRefresh" & vbCrLf & "No: Kill Auto Refresh" & vbCrLf & "Cancel: Exit Service", vbYesNoCancel + vbQuestion + vbDefaultButton1, "SmoothScroll AutoRefresh")

Select Case response
    Case vbYes
        MsgBox "Service Activated", vbInformation, "SmoothScroll AutoRefresh"
        batchScript = "@echo off" & vbCrLf & _
            ":LOOP" & vbCrLf & _
            "taskkill /F /IM SmoothScroll.exe" & vbCrLf & _
            "timeout /t 1 /nobreak >nul" & vbCrLf & _
            "rem Start SmoothScroll.exe" & vbCrLf & _
            "start """" """ & smoothScrollPath & """\SmoothScroll.exe""" & vbCrLf & _
            "timeout /t 600 /nobreak >nul" & vbCrLf & _
            "goto LOOP"

        tempBatchFile = objShell.ExpandEnvironmentStrings("%TEMP%\TempBatchFile.bat")
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objFile = objFSO.CreateTextFile(tempBatchFile, True)
        objFile.Write batchScript
        objFile.Close
        objShell.Run "cmd /k " & tempBatchFile, 0, True
        objFSO.DeleteFile(tempBatchFile)

    Case vbNo
        confirmResponse = MsgBox("Deactivating Service will kill cmd.exe and conhost.exe, Proceed?", vbYesNo + vbExclamation, "Confirmation")
        If confirmResponse = vbYes Then
            objShell.Run "taskkill /F /IM conhost.exe", 0, True
            objShell.Run "taskkill /F /IM cmd.exe", 0, True

            MsgBox "Service Deactivated", vbInformation, "SmoothScroll AutoRefresh"
        Else
        End If

    Case vbCancel
End Select

Function FindSmoothScrollPath()
    Dim objWMIService, colProcesses, objProcess
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'SmoothScroll.exe'")
    For Each objProcess In colProcesses
        FindSmoothScrollPath = Left(objProcess.ExecutablePath, InStrRev(objProcess.ExecutablePath, "\") - 1)
        Exit Function
    Next
    FindSmoothScrollPath = ""
End Function
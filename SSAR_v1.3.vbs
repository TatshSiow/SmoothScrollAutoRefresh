Set objShell = CreateObject("WScript.Shell")

' Find the location of SmoothScroll.exe
smoothScrollPath = FindSmoothScrollPath()

' Display a dialog box with "Run," "Terminate," and "Exit" buttons
response = MsgBox("Yes: Run AutoRefresh" & vbCrLf & "No: Kill Auto Refresh" & vbCrLf & "Cancel: Exit Service", vbYesNoCancel + vbQuestion + vbDefaultButton1, "SmoothScroll AutoRefresh")

' Choose an action based on the user's response
Select Case response
    Case vbYes ' Run
        ' Display a prompt indicating Service Activated
        MsgBox "Service Activated", vbInformation, "SmoothScroll AutoRefresh"
        ' Batch script content
        batchScript = "@echo off" & vbCrLf & _
            ":LOOP" & vbCrLf & _
            "taskkill /F /IM SmoothScroll.exe" & vbCrLf & _
            "timeout /t 1 /nobreak >nul" & vbCrLf & _
            "rem Start SmoothScroll.exe" & vbCrLf & _
            "start """" """ & smoothScrollPath & """\SmoothScroll.exe""" & vbCrLf & _
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
        ' Terminate the running batch script
        TerminateBatchScript

        ' Display a prompt indicating Service Deactivated
        MsgBox "Service Deactivated", vbInformation, "SmoothScroll AutoRefresh"

    Case vbCancel ' Exit
        ' Do nothing or add any cleanup logic if needed

End Select

Function FindSmoothScrollPath()
    ' Function to find the location of SmoothScroll.exe
    Dim objWMIService, colProcesses, objProcess

    ' Create WMI service object
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

    ' Query for SmoothScroll.exe process
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'SmoothScroll.exe'")

    ' Check if SmoothScroll.exe process is running
    For Each objProcess In colProcesses
        FindSmoothScrollPath = Left(objProcess.ExecutablePath, InStrRev(objProcess.ExecutablePath, "\") - 1)
        Exit Function
    Next

    ' Return an empty string if SmoothScroll.exe is not found
    FindSmoothScrollPath = ""
End Function

Sub TerminateBatchScript()
    ' Subroutine to terminate the running batch script
    Dim objWMIService, colProcesses, objProcess

    ' Create WMI service object
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

    ' Query for CMD.exe process associated with the batch script
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where CommandLine Like '%TempBatchFile.bat%'")

    ' Terminate the CMD.exe process associated with the batch script
    For Each objProcess In colProcesses
        objProcess.Terminate
    Next
End Sub

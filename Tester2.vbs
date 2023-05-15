Dim objShell, objFSO, outputPath, currentDate

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the directory where the script is located
scriptDirectory = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Change the current directory to the script directory
objShell.CurrentDirectory = scriptDirectory

' Show folder browse dialog using FileDialog object
Set folderDialog = CreateObject("Shell.Application").BrowseForFolder(0, "Select the directory to store files:", 0, scriptDirectory)

' Check if a folder was selected
If folderDialog Is Nothing Then
    MsgBox "No directory selected. Script will terminate."
    WScript.Quit
End If

' Get the selected folder path
outputPath = folderDialog.Self.Path

' Execute the command to retrieve the current date in the desired format
currentDate = objShell.Exec("%COMSPEC% /c echo %date:~10,4%%date:~4,2%%date:~7,2%").StdOut.ReadAll

' Remove any line breaks or carriage returns from the current date
currentDate = Replace(Replace(currentDate, vbCrLf, ""), vbCr, "")

' Change the current directory back to the script directory
objShell.CurrentDirectory = scriptDirectory

' Display the directory tree
objShell.Run "%COMSPEC% /c tree " & Chr(34) & outputPath & Chr(34) & " /f", 1, True

' Execute commands and write output files
PerformCommand "ipconfig /all", "IP Check"
PerformCommand "systeminfo", "System Info"
PerformCommand "tasklist", "Task List"

' Keep the command prompt window open
objShell.Run "%COMSPEC% /k", 1, False

' Function to execute a command and write the output to a file
Sub PerformCommand(command, commandType)
    Dim objExec, output, fileName, file, invalidChars, i

    ' Define characters not allowed in filenames
    invalidChars = Array("<", ">", ":", """", "/", "\", "|", "?", "*")

    ' Generate the filename based on the formatted date, computer name, user name, and command type
    fileName = currentDate & "-" & objShell.ExpandEnvironmentStrings("%COMPUTERNAME%") & "-" & objShell.ExpandEnvironmentStrings("%USERNAME%") & "-" & commandType & ".txt"

    ' Remove invalid characters from the filename
    For i = 0 To UBound(invalidChars)
        fileName = Replace(fileName, invalidChars(i), "")
    Next

    ' Create the output file
    Set file = objFSO.CreateTextFile(objFSO.BuildPath(outputPath, fileName), True)

    ' Execute the command and write the output to the file
    Set objExec = objShell.Exec(command)
    
    ' Display the heading of the process
    WScript.StdOut.WriteLine "Executing: " & command & vbCrLf

    ' Read the command output and write it to the file
    Do While Not objExec.StdOut.AtEndOfStream
        output = objExec.StdOut.ReadLine
        file.WriteLine output
        
        ' Display the result line by line
        WScript.StdOut.WriteLine output
    Loop

    ' Close the output file
    file.Close
    
    ' Add a line break after each process
    WScript.StdOut.WriteLine
End Sub

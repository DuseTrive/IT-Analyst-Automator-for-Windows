Dim objShell, objFSO, outputPath, currentDate, objExec
' paddings
topSpacing = String(5, vbCrLf) ' Add 5 lines of spacing at the top
middleSpacing = String(1, vbCrLf) ' Add 5 lines of spacing at the top
bottomSpacing = String(5, vbCrLf) ' Add 5 lines of spacing at the bottom

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Display ASCII art when the program starts
WScript.Echo "+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+"
WScript.Echo "|  /$$$$$$                        /$$                                   /$$$$$$            /$$$$$$                /$$$$$$$                                            /$$                        |"
WScript.Echo "| /$$__  $$                      | $$                                  |_  $$_/           /$$__  $$              | $$__  $$                                          | $$                        |"
WScript.Echo "|| $$  \__/ /$$   /$$  /$$$$$$$ /$$$$$$    /$$$$$$  /$$$$$$/$$$$         | $$   /$$$$$$$ | $$  \__//$$$$$$       | $$  \ $$  /$$$$$$   /$$$$$$   /$$$$$$   /$$$$$$  /$$$$$$    /$$$$$$   /$$$$$$ |"
WScript.Echo "||  $$$$$$ | $$  | $$ /$$_____/|_  $$_/   /$$__  $$| $$_  $$_  $$        | $$  | $$__  $$| $$$$   /$$__  $$      | $$$$$$$/ /$$__  $$ /$$__  $$ /$$__  $$ /$$__  $$|_  $$_/   /$$__  $$ /$$__  $$|"
WScript.Echo "| \____  $$| $$  | $$|  $$$$$$   | $$    | $$$$$$$$| $$ \ $$ \ $$        | $$  | $$  \ $$| $$_/  | $$  \ $$      | $$__  $$| $$$$$$$$| $$  \ $$| $$  \ $$| $$  \__/  | $$    | $$$$$$$$| $$  \__/|"
WScript.Echo "| /$$  \ $$| $$  | $$ \____  $$  | $$ /$$| $$_____/| $$ | $$ | $$        | $$  | $$  | $$| $$    | $$  | $$      | $$  \ $$| $$_____/| $$  | $$| $$  | $$| $$        | $$ /$$| $$_____/| $$      |"
WScript.Echo "||  $$$$$$/|  $$$$$$$ /$$$$$$$/  |  $$$$/|  $$$$$$$| $$ | $$ | $$       /$$$$$$| $$  | $$| $$    |  $$$$$$/      | $$  | $$|  $$$$$$$| $$$$$$$/|  $$$$$$/| $$        |  $$$$/|  $$$$$$$| $$      |"
WScript.Echo "| \______/  \____  $$|_______/    \___/   \_______/|__/ |__/ |__/      |______/|__/  |__/|__/     \______/       |__/  |__/ \_______/| $$____/  \______/ |__/         \___/   \_______/|__/      |"
WScript.Echo "|           /$$  | $$                                                                                                                | $$                                                        |"
WScript.Echo "|          |  $$$$$$/                                                                                                                | $$                                                        |"
WScript.Echo "|           \______/                                                                                                                 |__/                                                        |"
WScript.Echo "|++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++|"
WScript.Echo "|Creator: DuseTrive - https://github.com/DuseTrive                                                                                                                                               |"
WScript.Echo "|License: MIT License                                                                                                                                                                            |"
WScript.Echo "|Description: This program is to automate CPU and System checkups. For More information check on my GitHub Page                                                                                  |"
WScript.Echo "+------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
WScript.Sleep 3000 ' Sleep for 3 seconds (adjust as needed)
WScript.Echo topSpacing
WScript.Echo "Select the Folder to save the results in popup window"
WScript.Echo bottomSpacing
WScript.Sleep 3000 ' Sleep for 3 seconds (adjust as needed)

Do
    ' Show folder browse dialog using FileDialog object
    Set folderDialog = CreateObject("Shell.Application").BrowseForFolder(0, "Select the directory to store files:", &H4000)

    ' Check if a folder was selected
    If folderDialog Is Nothing Then
        ' Prompt to confirm quitting the program

        Dim answer
        answer = MsgBox("No directory selected. Do you want to quit the program?", vbYesNo + vbExclamation, "Warning")

        If answer = vbYes Then
            ' User confirmed to quit the program
            MsgBox "The programme is quite", vbCritical, "Error"
            WScript.Echo "The programme is quite"
            WScript.Quit
        End If
    Else
        ' Folder selected, display the folder path and ask for confirmation
        Dim folderPath
        folderPath = folderDialog.Self.Path
        Dim confirmAnswer
        confirmAnswer = MsgBox("Selected folder: " & folderPath & vbCrLf & vbCrLf & "Do you want to save files in this folder?", vbYesNo + vbQuestion, "Confirmation")
        
        If confirmAnswer = vbYes Then
            ' User confirmed to save files in the selected folder, continue with the script
            MsgBox "Files will be saved in this location:" & vbCrLf & folderPath, vbExclamation, "Warning"
            WScript.Echo "Folder location:" & folderPath
            Exit Do
        End If
    End If
Loop

' Rest of the code...
outputPath = folderPath


' Get the selected folder path
outputPath = folderDialog.Self.Path

' Rest of the code...


' Get the selected folder path
outputPath = folderDialog.Self.Path

' Execute the command to retrieve the current date in the desired format
currentDate = objShell.Exec("%COMSPEC% /c echo %date:~10,4%%date:~4,2%%date:~7,2%").StdOut.ReadAll

' Remove any line breaks or carriage returns from the current date
currentDate = Replace(Replace(currentDate, vbCrLf, ""), vbCr, "")

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' You are free to Add any CMD Command below in order to Customize the information you need to Gather
' REveiw Commands.txt for more CMD Commands
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' Example useage:::::
' PerformCommand "CMD-Command", "file name"
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' Tempatate  to use
'PerformCommand "", ""
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

' Executable Commands
PerformCommand "ipconfig /all", "IP Check"
PerformCommand "systeminfo", "System Info"

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

' Function to execute a command and write the output to a file
Sub PerformCommand(command, commandType)
    Dim objExec, output, fileName, file, invalidChars, i
    Dim heading, headingLength, equalSigns, equalSignsLength, spacing, topSpacing, bottomSpacing

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

    ' Set the heading and format it for display
    heading = "=== " & commandType & " ==="
    headingLength = Len(heading)
    equalSignsLength = 40 ' Set the length of equal signs (adjust as needed)
    equalSigns = String(equalSignsLength, "=")
    spacing = String((equalSignsLength - headingLength) \ 2, " ")
    
    ' Display the heading in the CMD window
    WScript.Echo topSpacing
    WScript.Echo equalSigns
    WScript.Echo spacing & heading & spacing
    WScript.Echo equalSigns
    WScript.Echo bottomSpacing

    ' Write the heading to the file
    file.WriteLine topSpacing
    file.WriteLine equalSigns
    file.WriteLine spacing & heading & spacing
    file.WriteLine equalSigns
    file.WriteLine bottomSpacing

    ' Execute the command and write the output to the file and CMD window
    Set objExec = objShell.Exec(command)
    Do While Not objExec.StdOut.AtEndOfStream
        output = objExec.StdOut.ReadLine
        file.WriteLine output
        WScript.Echo output ' Display the output in the CMD window
    Loop

    ' Close the output file
    file.Close

End Sub

' Last guidens for user
WScript.Echo topSpacing
WScript.Echo "=========================================================================================================================================================================================="
WScript.Echo middleSpacing
WScript.Echo "All the text file are saved in this Folder location: " & folderPath
WScript.Echo middleSpacing
WScript.Echo "!!!!!!!!!!!!!!!!!!!!  Please note that if the saved files do not have valid output, you may need to rerun the program with administrator privileges. !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
WScript.Echo bottomSpacing

' Display end of program
MsgBox "Hurray! System Info gathering is over.", vbExclamation, "End of Program"

' Display folder location to view files
Dim FinAnswer
FinAnswer = MsgBox("Files have been saved in the following location:" & vbCrLf & outputPath & vbCrLf & vbCrLf & "Do you want to open the folder location?", vbQuestion + vbYesNo, "File Location")

If FinAnswer = vbYes Then
    ' Open folder location
    objShell.Run "explorer.exe " & outputPath
End If

' End of the script
WScript.Quit

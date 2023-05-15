Dim objShell, objFSO, outputPath, currentDate

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Show folder browse dialog using FileDialog object
Set folderDialog = CreateObject("Shell.Application").BrowseForFolder(0, "Select the directory to store files:", &H4000)

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

' Execute commands and write output files
PerformCommand "ipconfig /all", "IP Check"
PerformCommand "systeminfo", "System Info"
PerformCommand "wmic cpu get Name, Manufacturer, MaxClockSpeed, NumberOfCores", "CPU Info"
PerformCommand "wmic path Win32_VideoController get Name, AdapterCompatibility, DriverVersion", "GPU Info"
PerformCommand "wmic path Win32_OperatingSystem get Name, Caption, OSArchitecture, InstallDate, SerialNumber", "Operating System Info"
PerformCommand "wmic path Win32_Product get Name, Version, Vendor, InstallDate", "Installed Software Info"
PerformCommand "sfc /scannow", "System File Checker"
PerformCommand "chkdsk /f /r", "Disk Check"
PerformCommand "wmic path Win32_PerfFormattedData_PerfOS_Memory get CommittedBytes, FreeAndZeroPageListBytes, CacheBytes, PoolPagedBytes, PoolNonpagedBytes", "Memory Info"
PerformCommand "wmic path Win32_PerfFormattedData_PerfOS_Processor get Name, PercentProcessorTime", "Processor Usage"
PerformCommand "wmic path Win32_PerfFormattedData_PerfDisk_PhysicalDisk get Name, AvgDiskQueueLength, PercentDiskTime", "Disk Usage"
PerformCommand "wmic path Win32_PerfFormattedData_Tcpip_NetworkInterface get Name, BytesTotalPersec, BytesReceivedPersec, BytesSentPersec", "Network Usage"
PerformCommand "wmic path Win32_PerfFormattedData_PerfOS_System get ProcessorQueueLength", "System Queue Length"
PerformCommand "wmic path Win32_PerfFormattedData_PerfOS_PagingFile get Name, PercentUsage", "Page File Usage"
PerformCommand "wmic path SoftwareLicensingService get OA3xOriginalProductKey", "Windows License Key"


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
    Do While Not objExec.StdOut.AtEndOfStream
        output = objExec.StdOut.ReadLine
        file.WriteLine output
    Loop

    ' Close the output file
    file.Close
End Sub

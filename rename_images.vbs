''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       SCRIPT TO RENAME MULTIPLE FILES              '
'           MADE IN 2022/05/02                       '
'           EVERTON SILVA                            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''
' How it works?
' 1. Put the executable file in a different folder than the folder you want to rename the files.
'
' 2. Run this file: rename_images.vbs
' 
' 3. Paste the full path of the folder C:\Users\Example\Documents\Images
' 
' 4. Then choose a new prefix for the files. The files will look like this: example_01 example_02 example_03
'
' 5. At the end of the execution, the number of renamed and unrenamed files will be displayed.
''''''''''''''''''''''''''''''''''''

' Variables to run the script
StartExecutionTime = Now
Set FSO = CreateObject("Scripting.FileSystemObject")

LogCopyFiles = ""
totalRenamedFiles = 0
LogCopyError = ""
totalCopyError = 0
LogCopyDup = ""
totalCopyDup = 0
tempName = ""

Dim FolderPath
FolderPath = InputBox("Paste the full path of the folder.")
If FolderPath = "" Then
    MsgBox("See you soon")
    Wscript.Quit
End If

Do Until FSO.FolderExists(FolderPath)
    FolderPath = InputBox("Paste the full path of the folder.")
    If FolderPath = "" Then
        MsgBox("See you soon")
        Wscript.Quit
    End If
Loop
scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

Do While FolderPath = scriptdir
    MsgBox("Please put the executable file in a different folder")
    FolderPath = InputBox("Paste the full path of the folder.")
        If FolderPath = "" Then
            MsgBox("See you soon")
            Wscript.Quit
        End If
Loop

Dim PrefixName
PrefixName = InputBox("Choose a new prefix for the files.")
If PrefixName = "" Then
    MsgBox("See you soon")
    Wscript.Quit
End If

PrefixName =  Replace(PrefixName," ","-",1,-1)
PrefixName =  LCase(PrefixName)

' CLEAR LOG FILE
if FSO.FileExists("files.log") then
	FSO.DeleteFile("files.log")
end if

Dim iCounter
Dim folder, file

' Path
Set folder = fso.GetFolder(FolderPath)  

iCounter = 1
For Each file In folder.Files
    ' get the file extension
    extension = FSO.GetExtensionName(file.Name)
    if extension = "" Then
        extension = "jpeg"
    end if
    tempName = PrefixName & "_" & iCounter & "." & extension
    FileFullPath = FolderPath & "\" & tempName
    if FSO.FileExists(FileFullPath) Then
        'FILE ALREADY EXISTS			
        LogCopyDup = LogCopyDup & vbCrlf &  FileFullPath
        totalCopyDup = totalCopyDup + 1
    else 
        file.Name =  tempName
        totalRenamedFiles = totalRenamedFiles + 1
        LogCopyFiles = LogCopyFiles & vbCrlf &  FileFullPath
    end if
    iCounter = iCounter + 1
Next

EndExecutionTime = Now
TempoExec = DateDiff("s", StartExecutionTime, EndExecutionTime)

' WRITE LOG
totalNotRenamed = totalCopyDup

LogCopyTxt = "Start: " & StartExecutionTime & vbCrlf & "End: " & EndExecutionTime & vbCrlf & TempoExec & " seconds of execution!" & vbCrlf & totalRenamedFiles & " renamed files!" & vbCrlf & totalNotRenamed & " unrenamed files!" & vbCrlf & vbCrlf & vbCrlf & "===========RENAMED FILES (" & totalRenamedFiles & ")===========" & vbCrlf & LogCopyFiles & vbCrlf & vbCrlf & vbCrlf & "===========FILES ALREADY EXISTS (" & totalCopyDup & ")===========" & vbCrlf & LogCopyDup & vbCrlf & vbCrlf & vbCrlf & "===========FILES ERRORS (" & totalCopyError & ")===========" & vbCrlf & LogCopyError
set LogFile = FSO.CreateTextFile("arquivos.log", True)
LogFile.WriteLine(LogCopyTxt)
LogFile.Close

MsgBox("Finished!" & vbCrlf & totalRenamedFiles & " renamed files!" & vbCrlf & totalNotRenamed & " unrenamed files!" & vbCrlf & TempoExec & " seconds of execution!")
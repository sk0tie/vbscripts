'swxLogToFile.vbs
'Written by: Scott Wilcox
'Date: 01/13/2011
'The Cobalt Group, Inc.
'

'This script is a quick and dirty method to easily log messages to a log file.
'To log to a file, simply pass two arguments to this script; the log file name
'and the message being logged. The script will create the file if it doesn't
'exist, open the file and append the message.  After each logged line, the file
'is closed, so the log file can be called by multiple sources if needed.
'

'Example Usage: swxLogToFile.vbs "example.log" "This is an example log file."
'

'If you're using VBScript to execute this, the following code snippet can be used
'to quickly setup logging within that script. Modify the strLogScript with the
'location of this script and strLogName with the location of the new log file.
'
'Example VBScript Method:
'
'Sub LogToFile(strLogMessage)
'	Dim objShell
'	Dim strLogScript, strLogName, strCommand
'
'	strLogScript = "swxLogToFile.vbs"
'	strLogName = "example.log"
'
'	strCommand = strLogScript & " " & Chr(34) & strLogName & Chr(34) & " " _
' 	& Chr(34) & Now & " " & strLogMessage & Chr(34)
'	
'	Set objShell = CreateObject("Wscript.Shell")
'	objShell.Run "wscript " & strCommand, , True
'
'	Set objShell = Nothing
'End Sub
'

'Use this option to make sure that every variable has been declared and no random
'variables are created by accident.
'
Option Explicit

'Declare our variables; I like to keep my variable types on seperate lines.
'
Dim objFSO, objLog
Dim strLogName, strLogMessage

'Make sure the proper number of arguments were passed to the script, otherwise
'exit the script gracefully.
'
If (Wscript.Arguments.Count = 2) Then
	strLogName = Wscript.Arguments(0)
	strLogMessage = Wscript.Arguments(1)
Else
	Wscript.Echo "Requires the log file to append and the log message to be written." _
	& vbCrLF & vbCrLf & "Example: swxLogToFile.vbs ""example.log"" ""This is an example log file."""
	Wscript.Quit(0)
End If

'Create our file object handler and create/open the file specified for appending.
'
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLog = objFSO.OpenTextFile(strLogName, 8, True)

'Write the message passed to the script to the log file and close the file.
'
objLog.WriteLine(strLogMessage)
objLog.Close

'Clean up our objects as VBscript is bad at garbage collection.
'
Set objFSO = Nothing
Set objLog = Nothing
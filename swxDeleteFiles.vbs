'swxDeleteFiles.vbs
'Written by: Scott Wilcox
'Date: 01/13/2011
'The Cobalt Group, Inc.
'
'This script facilitates directory cleanup and can be run from one host and scheduled to delete
'files locally or through UNC paths on remote hosts.  If the script is used to cleanup directories on
'remote hosts, it must be run or scheduled by a user with administrative rights on the remote hosts.
'Scheduling the script can be done a variety of ways and is out of scope for this script.  The most
'common method would be to use Windows Scheduled Tasks. The script can also log to a file if needed
'by adding a fourth argument with the location of the log file to create and append to, however
'this does reduce performance drastically.
'

'Use this option to make sure that every variable has been declared and no random
'variables are created by accident.
'
Option Explicit

'Upon error, resume if not fatal and log the error..
'

'Declare our variables; I like to keep my variable types on seperate lines.
'
Dim objFSO, objDirectory, objFile
Dim strDirectory, strInterval, strLogName
Dim intRetention, intAge
Dim arrFiles
Dim boolLog

'Populate variables with the arguments passed when running the script, making sure that they're explicitly
'defined as their proper types. Make sure that the script was called with the number of arguments we want
'to use, if not, explain the basics of running the script.
'

If (Wscript.Arguments.Count >= 3) Then
	strDirectory = CStr(Wscript.Arguments(0))
	intRetention = CInt(Wscript.Arguments(1))
	strInterval = CStr(Wscript.Arguments(2))
	boolLog = False
	If (Wscript.Arguments.Count = 4) Then
		strLogName = CStr(Wscript.Arguments(3))
		boolLog = True
	End If
Else
	Wscript.Echo "Requires directory path, retention time interval and optional log file name." _
	& vbCrLf & vbCrLf & "Example: swxDeleteFiles.vbs ""\\examplehost\c$\logfiles\"" 30 d ""debug.log""" _
	& vbCrLf & vbCrLf & "Valid Intervals include:" & vbCrLf & vbCrLf & vbTab & "yyyy (Year)" _
	& vbCrLf & vbTab & "q (Quarter)" & vbCrLf & vbTab & "m (Month)" & vbCrLf & vbTab & "y (Day of Year)" _
	& vbCrLf & vbTab & "d (Day)" & vbCrLf & vbTab & "w (Weekday)" & vbCrLf & vbTab & "ww (Week of Year)" _
	& vbCrLf & vbTab & "h (Hour)" & vbCrLf & vbTab & "m (Minute)" & vbCrLf & vbTab & "s (Second)"
	Wscript.Quit(0)
End If

'Begin error handling and build our system objects that handle the files we'll be checking.
'

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objDirectory = objFSO.GetFolder(strDirectory)
Set arrFiles = objDirectory.Files

'Write to the log that we're beginning the process.
'
LogToFile("Directory clean up started for " & strDirectory)
LogToFile("Deleting files older than " & intRetention & strInterval & ".")

For Each objFile in arrFiles
	intAge = DateDiff(strInterval, objFile.DateCreated, Now)
	If (intAge > intRetention) Then
		LogToFile(vbTab & "Deleting " & objFile & " Age:" & intAge & strInterval)
		'objFile.Delete
	End If
Next

'Write to the log, that we're finished with the process.
'
LogToFile("Directory clean up completed for " & strDirectory)

'Clean up our objects as VBscript is bad at garbage collection.
'
Set objDirectory = Nothing
Set objFile = Nothing
Set objFSO = Nothing

'Logging method. Refer to swxLogToFile.vbs for more information.
'
Sub LogToFile(strLogMessage)
	'If logging is enabled, log to the specified file..
	'
	If (boolLog = True) Then

		Dim objShell
		Dim strLogScript, strCommand
	
		strLogScript = "swxLogToFile.vbs"
		strCommand = strLogScript & " " & Chr(34) & strLogName & Chr(34) & " " & Chr(34) & Now & " " & strLogMessage & Chr(34)

		Set objShell = CreateObject("Wscript.Shell")
		objShell.Run "wscript " & strCommand, , True

		Set objShell = Nothing
	End If
End Sub
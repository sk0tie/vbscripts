Option Explicit

Dim strMemberName, strMemberPath, strProxyAddress
Dim intMemberCount, intAddressCount
Dim arrProxyAddresses
Dim objOU, objMember, objFSO, objLog

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLog = objFSO.OpenTextFile("swxProxyAddresses.log", 2, True)

Set objOU = GetObject("LDAP://cn=Microsoft Exchange System Objects,dc=main,dc=cobaltgroup,dc=com")

objLog.WriteLine(Now & " Starting ProxyAddress Dump for Public Folders")

For Each objMember in objOU
  If objMember.Class = "publicFolder" Then
  	strMemberName = objMember.DisplayName
	strMemberPath = objMember.FolderPathName
  	intMemberCount = intMemberCount + 1
  	arrProxyAddresses = objMember.ProxyAddresses
  	For Each strProxyAddress in arrProxyAddresses
		intAddressCount = intAddressCount + 1
  		objLog.WriteLine(strMemberName & vbTab & strMemberPath & vbTab & strProxyAddress)
		'objLog.WriteLine(strProxyAddress)
  	Next
  End If
Next

objLog.WriteLine(Now & " Completed ProxyAddress Dump of " & intAddressCount & " addresses for " & intMemberCount & " Public Folders")
objLog.Close

Set objFSO = Nothing
Set objLog = Nothing
Set objOU = Nothing
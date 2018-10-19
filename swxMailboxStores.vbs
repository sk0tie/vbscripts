'swxMailboxStores.vbs
'Written by: Scott Wilcox
'Date: 12/03/2009
'The Cobalt Group, Inc.

'VBScript using Collaboration Data Objects (CDO). The system that runs this script
'must have Mapi32.dll and cdo.dll from Exchange 5.5 or newer registered.
'This will open a connection to each Exchange Server, it's Storage Groups, thier
'Mailbox Stores and check their status, sending an alert if any stores are offline.
'

'Use this option to make sure that every variable has been declared and no random
'variables are created by accident.
'
Option Explicit

'Declare our variables; I like to keep my variable types on seperate lines.
'
Dim objServer, objStorageGroup, objMailboxStore, objMessage
Dim arrServers, arrStorageGroups, arrMailboxStores
Dim strServer, strStorageGroup, strStorageGroupURL, strMailboxStore, strMailboxStoreURL, strAlertSubject, strAlertBody
Dim intServer, intStorageGroup, intMailboxStore, intAlerts

'Setup our CDO objects to talk with Exchange.
'
Set objServer = CreateObject("CDOEXM.ExchangeServer")
Set objStorageGroup = CreateObject("CDOEXM.StorageGroup")
Set objMailboxStore = CreateObject("CDOEXM.MailboxStoreDB")

'This is the list of Exchange Virtual Servers (we check the virtuals, not the actual servers).
'If any new virtuals are added in the future, simply add an element to this array.
'
arrServers = array("SEA-EVS-01", "SEA-EVS-02", "LYN-EVS-01")

'These nested for loops will, for each server, open a data connection to each storage group
'and for each mailbox store in each storage group, pull out the names, and check their status.
'Status levels are as follows: 0 - Online, 1 - Offline, 2 - Mounting, 3 - Dismounting.
'Since we only care about the stores being online or offline, we can simply check if each
'stores status is greater than 0. If it is, we start building an alert message.
'Since there are so many stores, we don't want to send an e-mail for each store that goes
'offline, but rather one e-mail with a list of each store that is offline.
'
For intServer = 0 To UBound(arrServers)
	objServer.DataSource.Open arrServers(intServer)
	arrStorageGroups = objServer.StorageGroups
	strServer = arrServers(intServer)

	'Get the LDAP URL, open a data source and get the name of this storage group.
	'
	For intStorageGroup = 0 To UBound(arrStorageGroups)
		strStorageGroupURL = arrStorageGroups(intStorageGroup)
		objStorageGroup.DataSource.Open "LDAP://" & objServer.DirectoryServer & "/" & strStorageGroupURL
		strStorageGroup = objStorageGroup.Name
		arrMailboxStores = objStorageGroup.MailboxStoreDBs

		'Get the LDAP URL, open a data source and get the name of this store.
		'
		For intMailboxStore = 0 To UBound(arrMailboxStores)
			strMailboxStoreURL = arrMailboxStores(intMailboxStore)
			objMailboxStore.DataSource.Open "LDAP://" & strMailboxStoreURL
			strMailboxStore = objMailboxStore.Name

			'Check this store for it's status; if it's offline, increment our alert counter
			'and append a string to our alert subject.
			'			
			If objMailboxStore.Status > 0 Then
				intAlerts = intAlerts + 1
				strAlertSubject = "Exchange Mailbox Stores (" & intAlerts & " Offline)"
				strAlertBody = strAlertBody & "Server: " & strServer & " | Storage Group: " & strStorageGroup & " | Mailbox Store: " & strMailboxStore & " | Status: Offline" & vbCrLf
			End If
		Next
	Next
Next

'Check to see if we have any alerts; if we do, call the SendAlert Function.
'
If intAlerts > 0 Then
	SendAlert strAlertSubject, strAlertBody
End If

'Build our CDO messaging object and send our alert.
'
Function SendAlert(strAlertSubject, strAlertBody)
  Set objMessage = CreateObject("CDO.Message")
  objMessage.To = "corpitpager@cobaltgroup.com"
  objMessage.From = "ITAlerts@cobaltgroup.com"
  objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
  objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "franklin.cobaltgroup.com" 
  objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
  objMessage.Configuration.Fields.Update
  objMessage.Subject = 	strAlertSubject
  objMessage.TextBody = strAlertBody
  objMessage.Send
End Function

'Clean up our objects, since VBscript is bad at garbage collection.
'
Set objServer = Nothing
Set objStorageGroup = Nothing
Set objMailboxStore = Nothing
Set objMessage = Nothing
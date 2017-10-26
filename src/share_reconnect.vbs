' Network Share Reconnecter
'
' Purpose: This script tries to automatically reconnect
'          disconnected Windows network shares / drives 
'          if they are offline or are listed as offline.
'          The current network and access state is 
'          periodically checked until they become available 
'          or if the reconnection threshold is hit.
'
' Author: thexmanxy (Andreas Kar)
' Contact: andreas.kar@gmx.at
'----------------------------------------------------------

sub checkNconnect(hostname, sharePaths, shareLetters, sleep, pingCtn, netUseCtn, netUsePersistent)
	offline = 1
	i = 0
	set WshShell = CREATEOBJECT("WScript.Shell")
	while offline = 1 And i <= pingCtn - 1
		offline = WshShell.Run("ping -n 1 " & hostname, 0, True)
		i = i + 1
		If offline = 1 Then
			WScript.Sleep sleep
		End If
	wend
	status = 1
	i = 0
	ctn = 0
	while status = 1 And i <= netUseCtn - 1
		For j = 0 to uBound(sharePaths)
			If(ctn = 0) Then
				status = WshShell.Run("net use " & shareLetters(j) & " \\" & hostname & "\" & sharePaths(j) & " /persistent:" & netUsePersistent, 0, True)
			Else
				WshShell.Run "net use " & shareLetters(j) & " \\" & hostname & "\" & sharePaths(j) & " /persistent:" & netUsePersistent, 0, True
			End If
			ctn = ctn + 1
		next
		i = i + 1
		If status = 1 Then
			WScript.Sleep sleep
		End If
	wend
end sub

' hostname - IP or hostname of server (must be modified)
' sharePaths - all share paths on that server (must be modified)
' shareLetters - the share / drive letters for the defined paths (must be modified)
' pingWait - wait time after failed server ping
' reconWait - wait time after failed availability check
' pingCtn - how many pings per reconnect should be executed before giving up
' netUseCtn - how many net use fails per reconnect are allowed before giving up
' serverRetryCtn - how many overall reconnection tries should be executed
' netUsePersistent - should net use create a persistent share (yes/no)
' debug - enable or disable debug dialogs on current reconnection state
'---------------------------------------------------------------------------------
sub waitOnServerConnect(hostname, sharePaths, shareLetters, pingWait, reconWait, pingCtn, netUseCtn, serverRetryCtn, netUsePersistent, debug)
	i = 0
	path = "\\" & hostname & "\" & sharePaths(0)
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	online = objFSO.FolderExists(path)

	checkNconnect hostname, sharePaths, shareLetters, pingWait, pingCtn, netUseCtn, netUsePersistent
	while i <= serverRetryCtn - 1 And Not online
		i = i + 1
		If Not online Then
			If debug = 1 Then
				MsgBox("Server still offline, keep waiting...")
			End If
			WScript.Sleep reconWait
			online = objFSO.FolderExists(path)
		End If
		checkNconnect hostname, sharePaths, shareLetters, pingWait, pingCtn, netUseCtn, netUsePersistent
	wend
	If debug = 1 And Not online Then
		MsgBox("Server now online, drives reconnected!")
	End If
end sub

' CHANGE THE "hostname", "sharePaths" AND "shareLetters" BELOW (mandatory)!
' All other parameter changes are optional and depend on timeouts and favor.
'-------------------------------------------------
waitOnServerConnect "192.168.1.1", Array("path\to\share1", "path\to\share2"), Array("Z:", "Y:"), 5000, 10000, 5, 5, 15, "yes", 0

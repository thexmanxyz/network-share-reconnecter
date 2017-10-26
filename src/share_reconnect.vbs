'-----------------------------------------------------------'
' Network Share Reconnecter                                 '
'                                                           '
' Purpose: This script tries to automatically reconnect     '
'          disconnected Windows network shares / drives     '
'          if they are offline or are listed as offline.    '
'          The current network and access state is          '
'          periodically checked until they become available '
'          or if the reconnection threshold is hit.         '
'                                                           '
' Author: thexmanxy (Andreas Kar)                           '
' Contact: andreas.kar@gmx.at                               '
'-----------------------------------------------------------'

'----------------------------------------'
' Simple class for server configuration. '
'----------------------------------------'

Class ServerConfiguration
   Public hostname
   Public sharePaths
   Public shareLetters
   Public netUsePersistent
   Public hPath
   Public online
End Class

'----------------------------------------'
' Simple class for script configuration. '
'----------------------------------------'

Class ScriptConfiguration
   Public pingWait
   Public netUseWait
   Public reconWait
   Public pingCtn
   Public netUseCtn
   Public serverRetryCtn
   Public debug
End Class

'-------------------------------------------------------------------------------------------'
' Routine to check server connectivity and try to reconnect shares if the server is online. '
'                                                                                           '
' srvConfig - configuration object of the server                                            '
' scriptConfig - object for the global script configuration                                 '
'-------------------------------------------------------------------------------------------'

sub checkNconnect(srvConfig, scriptConfig)
	set WshShell = CREATEOBJECT("WScript.Shell")
	dim i, offline, status, ctn, netUseCmd
	
	i = 0
	offline = 1
	while offline = 1 And i <= scriptConfig.pingCtn - 1
		offline = WshShell.Run("ping -n 1 " & srvConfig.hostname, 0, True)
		i = i + 1
		If offline = 1 Then
			WScript.Sleep scriptConfig.pingWait
		End If
	wend
	
	i = 0
	ctn = 0
	status = 1
	while status = 1 And i <= scriptConfig.netUseCtn - 1
		For j = 0 to uBound(srvConfig.sharePaths)
			netUseCmd = "net use " & srvConfig.shareLetters(j) & " \\" & srvConfig.hostname & "\" & srvConfig.sharePaths(j) & " /persistent:" & srvConfig.netUsePersistent
			If(ctn = 0) Then
				status = WshShell.Run(netUseCmd, 0, True)
			Else
				WshShell.Run netUseCmd, 0, True
			End If
			ctn = ctn + 1
		next
		i = i + 1
		If status = 1 Then
			WScript.Sleep scriptConfig.netUseWait
		End If
	wend
end sub

'----------------------------------------------------------------------'
' Routine to check share availability and initiate share reconnection. '
'                                                                      '
' srvConfig - configuration object of the server                       '
' scriptConfig - object for the global script configuration            '
'----------------------------------------------------------------------'

sub waitOnServerConnect(srvConfig, scriptConfig)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	dim i
		
	i = 0
	srvConfig.hPath = "\\" & srvConfig.hostname & "\" & srvConfig.sharePaths(0)
	srvConfig.online = objFSO.FolderExists(srvConfig.hPath)

	If srvConfig.online Then
		checkNconnect srvConfig, scriptConfig
	Else
		while i <= scriptConfig.serverRetryCtn - 1 And Not srvConfig.online
			i = i + 1
			If Not srvConfig.online Then
				If scriptConfig.debug = 1 Then
					MsgBox("Server still offline, keep waiting...")
				End If
				WScript.Sleep scriptConfig.reconWait
				srvConfig.online = objFSO.FolderExists(srvConfig.hPath)
				If srvConfig.online Then
					checkNconnect srvConfig, scriptConfig
				End If
			End If
		wend
	End If
	If scriptConfig.debug = 1 And srvConfig.online Then
		MsgBox("Server now online, drives reconnected!")
	End If
end sub

'-------------------------------------------------------------------------------'
' This are the OPTIONAL script parameters which can be adapted to TUNE the      '
' the script if it reconnects to slow or to MINIMIZE the overhead.              '
'                                                                               '
' pingWait - wait time after failed server ping                                 '
' netUseWait - wait time after failed net use                                   '
' reconWait - wait time after failed availability check                         '
' pingCtn - how many pings per reconnect should be executed before giving up    '
' netUseCtn - how many net use fails per reconnect are allowed before giving up '
' serverRetryCtn - how many overall reconnection tries should be executed       '
' debug - enable or disable debug dialogs on current reconnection state         '
'-------------------------------------------------------------------------------'

Set scriptConfig = new ScriptConfiguration
scriptConfig.pingWait = 5000
scriptConfig.netUseWait = 5000
scriptConfig.reconWait = 10000
scriptConfig.pingCtn = 5
scriptConfig.netUseCtn = 5
scriptConfig.serverRetryCtn = 15
scriptConfig.debug = 0

'-----------------------------------------------------------------------------------'
' Here are the parameters which MUST be changed for your server or remote share.    '
' If you want to use this script in your topology it's MANDATORY to change the      '
' values below APPROPRIATELY.                                                       '
'                                                                                   '
' hostname - IP or hostname of the remote server (must be modified)                 '
' sharePaths - all share paths on the server (must be modified)                     '
' shareLetters - the share / drive letters for the defined paths (must be modified) '
' netUsePersistent - should net use create a persistent share (yes/no)              '
'-----------------------------------------------------------------------------------'

Set srvConfig = New ServerConfiguration
srvConfig.hostname = "192.168.1.1"
srvConfig.sharePaths = Array("path\to\share1", "path\to\share2")
srvConfig.shareLetters = Array("Z:", "Y:")
srvConfig.netUsePersistent = "yes"

'--------------'
' Start Script '
'--------------'

waitOnServerConnect srvConfig, scriptConfig
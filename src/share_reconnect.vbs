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
   Public shell
   Public fso
   Public debug
End Class

'-------------------------------------------------------------'
' Routine to ping a server with a predfined configuration     '
'                                                             '
' scriptConfig - object for the global script configuration   '
' srvConfig - configuration object of the server              '
'-------------------------------------------------------------'

function pingServer(scriptConfig, srvConfig)
	pingServer = scriptConfig.shell.Run("ping -n 1 " & srvConfig.hostname, 0, True)
End Function

'-------------------------------------------------------------'
' Routine that tries to reach server with ping and failover   '
'                                                             '
' scriptConfig - object for the global script configuration   '
' srvConfig - configuration object of the server              '
'-------------------------------------------------------------'

Function pingReachServer(scriptConfig, srvConfig)
	Dim i, offline
	
	i = 0
	offline = 1
	While offline = 1 And i <= scriptConfig.pingCtn - 1
		offline = pingServer(scriptConfig, srvConfig)
		i = i + 1
		If offline = 1 Then
			WScript.Sleep scriptConfig.pingWait
		End If
	Wend
	pingReachServer = offline
	
End Function

'---------------------------------------------------------'
' Routine to create a net use command for a share         '
'                                                         '
' srvConfig - configuration object of the server          '
' pos - the current position of the share array           '
'---------------------------------------------------------'

Function createNetUseCmd(srvConfig, pos)
	createNetUseCmd = "net use " & srvConfig.shareLetters(pos) & " \\" & srvConfig.hostname & "\" & srvConfig.sharePaths(pos) & " /persistent:" & srvConfig.netUsePersistent
End Function

'---------------------------------------------------------------------'
' Routine that tries to reconnect shares with net uses and failover   '
'                                                                     '
' scriptConfig - object for the global script configuration           '
' srvConfig - configuration object of the server                      '
'---------------------------------------------------------------------'

Function netUseServerShares(scriptConfig, srvConfig)
	Dim i, status, ctn
	
	i = 0
	ctn = 0
	status = 1
	While status = 1 And i <= scriptConfig.netUseCtn - 1
		For j = 0 to uBound(srvConfig.sharePaths)
			If ctn = 0 Then
				status = scriptConfig.shell.Run(createNetUseCmd(srvConfig, j), 0, True)
			Else
				scriptConfig.shell.Run createNetUseCmd(srvConfig, j), 0, True
			End If
			ctn = ctn + 1
		Next
		i = i + 1
		If status = 1 Then
			WScript.Sleep scriptConfig.netUseWait
		End If
	Wend
End Function

'-------------------------------------------------------------'
' Routine to set server state by ping and share connectivity  '
'                                                             '
' scriptConfig - object for the global script configuration   '
' srvConfig - configuration object of the server              '
'-------------------------------------------------------------'

Sub setSrvState(ByVal scriptConfig, ByRef srvConfig)
	If pingReachServer(scriptConfig, srvConfig) = 0 Then
		srvConfig.online = scriptConfig.fso.FolderExists(srvConfig.hPath)
	Else
		srvConfig.online = 0
	End If
End Sub

'-------------------------------------------------------------'
' Routine to set server help path                             '
'                                                             '
' srvConfig - configuration object of the server              '
'-------------------------------------------------------------'

Sub setSrvPath(ByRef srvConfig)
	srvConfig.hPath = "\\" & srvConfig.hostname & "\" & srvConfig.sharePaths(0)
End Sub

'-------------------------------------------------------------'
' Routine that dynamically adjusts wait time on retry count   '
'                                                             '
' scriptConfig - object for the global script configuration   '
' retries - number of already executed retries                '
'-------------------------------------------------------------'

Function getReconWaitTime(scriptConfig, retries)
	Dim reconWait
	
	If retries <= 25 Then
		reconWait = scriptConfig.reconWait
	ElseIf retries > 25 and retries <= 40 Then
		reconWait = scriptConfig.reconWait * 6
	Else
		reconWait = scriptConfig.reconWait * 8
	End If
	
	getReconWaitTime = reconWait
End Function

'-------------------------------------------------------------------------------------------'
' Routine to check server connectivity and try to reconnect shares if the server is online. '
'                                                                                           '
' scriptConfig - object for the global script configuration                                 '
' srvConfig - configuration object of the server                                            '
'-------------------------------------------------------------------------------------------'

Sub checkNconnect(scriptConfig, srvConfig)
	status = pingReachServer(scriptConfig, srvConfig)
	If status = 0 Then
		netUseServerShares scriptConfig, srvConfig
	End If
End Sub

'----------------------------------------------------------------------'
' Routine to check share availability and initiate share reconnection. '
'                                                                      '
' scriptConfig - object for the global script configuration            '
' srvConfig - configuration object of the server                       '
'----------------------------------------------------------------------'

Sub waitOnServerConnect(scriptConfig, srvConfig)
	Set scriptConfig.fso = CreateObject("Scripting.FileSystemObject")
	Set scriptConfig.shell = CreateObject("WScript.Shell")
	Dim i
		
	i = 0
	setSrvPath srvConfig
	setSrvState scriptConfig, srvConfig
	While ((i <= scriptConfig.serverRetryCtn - 1 And Not srvConfig.online) Or (i = 0 And srvConfig.online))
		i = i + 1
		If Not srvConfig.online Then
			If scriptConfig.debug = 1 Then
				MsgBox("Server still offline, keep waiting...")
			End If
			
			WScript.Sleep getReconWaitTime(scriptConfig, i)
			setSrvState scriptConfig, srvConfig
			If srvConfig.online Then
				checkNconnect scriptConfig, srvConfig
			End If
		Else
			checkNconnect scriptConfig, srvConfig
		End If
	Wend
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
scriptConfig.pingWait = 250
scriptConfig.netUseWait = 250
scriptConfig.reconWait = 2500
scriptConfig.pingCtn = 3
scriptConfig.netUseCtn = 3
scriptConfig.serverRetryCtn = 75
scriptConfig.debug = 1

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

waitOnServerConnect scriptConfig, srvConfig
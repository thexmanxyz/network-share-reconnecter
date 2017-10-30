'--------------------------------------------------------------'
'   Network Share Reconnecter                                  '
'                                                              '
'   Purpose: This script tries to automatically reconnect      '
'            disconnected Windows network shares and drives    '
'            if they are offline or are listed as offline.     '
'            The current network and access state is           '
'            periodically checked until they become available  '
'            or when the reconnection threshold is hit.        '
'                                                              '
'   Author: Andreas Kar (thex) <andreas.kar@gmx.at >           '
'--------------------------------------------------------------'

'--------------------------------------------'
' Simple class for the server configuration. '
'--------------------------------------------'

Class ServerConfiguration
	Public hostname
	Public sharePaths
	Public shareLetters
	Public netUsePersistent
	Public hPath
	Public online
	Public connected
End Class

'--------------------------------------------'
' Simple class for the script configuration. '
'--------------------------------------------'

Class ScriptConfiguration
	Public pingWait
	Public netUseWait
	Public reconWait
	Public pingCtn
	Public netUseCtn
	Public serverRetryCtn
	Public pingTimeout
	Public winMgmts
	Public shell
	Public fso
	Public debug
End Class

'-----------------------------------------------------------------'
' Routine to Shell ping a server with a predfined configuration.  '
' (not used because ping does not return correct status)          '
'                                                                 '
' scriptConfig - object for the global script configuration       '
' srvConfig - configuration object of the server                  '
'-----------------------------------------------------------------'

Function pingServer(scriptConfig, srvConfig)
	pingServer = scriptConfig.shell.Run("ping -n 1 " & srvConfig.hostname, 0, True)
End Function

'-------------------------------------------------------------'
' Routine to create a WMI query for an ICMP ping.             '
'                                                             '
' scriptConfig - object for the global script configuration   '
' srvConfig - configuration object of the server              '
'-------------------------------------------------------------'

Function getWMIPingCmd(scriptConfig, srvConfig)
	getWMIPingCmd = "select * from Win32_PingStatus where TimeOut = " _ 
					& scriptConfig.pingTimeout & " and address = '" & srvConfig.hostname & "'"
End Function

'-----------------------------------------------------------------'
' Routine to ICMP ping a server with a predfined configuration.   '
'                                                                 '
' scriptConfig - object for the global script configuration       '
' srvConfig - configuration object of the server                  '
'-----------------------------------------------------------------'

Function pingICMPServer(scriptConfig, srvConfig)
	Dim ping, pEle, online
	
	online = false
	Set ping = scriptConfig.winMgmts.ExecQuery(getWMIPingCmd(scriptConfig, srvConfig))								
	For each pEle in ping
		online = Not IsNull(pEle.StatusCode) And pEle.StatusCode = 0
		If Not online Then
			Exit For
		End If
	Next
	pingICMPServer = Not online
End Function

'-----------------------------------------------------------------'
' Routine that retries to reach a server with a ping (failover).  '
'                                                                 '
' scriptConfig - object for the global script configuration       '
' srvConfig - configuration object of the server                  '
'-----------------------------------------------------------------'

Function retryPingServer(scriptConfig, srvConfig, icmp)
	Dim i, offline
	
	i = 0
	offline = true
	While offline And i <= scriptConfig.pingCtn - 1
		If icmp Then
			offline = pingICMPServer(scriptConfig, srvConfig)
		Else
			offline = pingServer(scriptConfig, srvConfig)
		End If
		i = i + 1
		If offline Then
			WScript.Sleep scriptConfig.pingWait
		End If
	Wend
	retryPingServer = offline
End Function

'------------------------------------------------------------'
' Routine to get the net use command for the current share.  '
'                                                            '
' srvConfig - configuration object of the server             '
' pos - the current position of the share array              '
'------------------------------------------------------------'

Function getNetUseCmd(srvConfig, pos)
	getNetUseCmd = "net use " & srvConfig.shareLetters(pos) & " \\" & srvConfig.hostname _ 
				   & "\" & srvConfig.sharePaths(pos) & " /persistent:" & srvConfig.netUsePersistent
End Function

'----------------------------------------------------------------------'
' Routine that tries to reconnect shares with net use and a failover.  '
'                                                                      '
' scriptConfig - object for the global script configuration            '
' srvConfig - configuration object of the server                       '
'----------------------------------------------------------------------'

Sub netUseServerShares(ByVal scriptConfig, ByRef srvConfig)
	Dim i, failed, ctn
	
	i = 0
	ctn = 0
	failed = 1
	While failed = 1 And i <= scriptConfig.netUseCtn - 1
		For j = 0 to uBound(srvConfig.sharePaths)
			If ctn = 0 Then
				failed = scriptConfig.shell.Run(getNetUseCmd(srvConfig, j), 0, True)
			Else
				scriptConfig.shell.Run getNetUseCmd(srvConfig, j), 0, True
			End If
			ctn = ctn + 1
		Next
		i = i + 1
		If failed = 1 Then
			WScript.Sleep scriptConfig.netUseWait
		End If
	Wend
	srvConfig.connected = false
End Sub

'------------------------------------------------------------------------'
' Routine to initialize the script config with objects that are reused.  '
'                                                                        '
' scriptConfig - object for the global script configuration              '
'------------------------------------------------------------------------'

Sub initConfig(ByRef scriptConfig)
	Set scriptConfig.winMgmts = GetObject("winmgmts:{impersonationLevel=impersonate}")
	Set scriptConfig.fso = CreateObject("Scripting.FileSystemObject")
	Set scriptConfig.shell = CreateObject("WScript.Shell")
End Sub

'--------------------------------------------------------------------'
' Routine that creates a new server object.                          '
'                                                                    '
' hostname - IP or human readable hostname                           '
' sharePaths - array that contains share paths                       '
' shareLetters - array that contains share letters                   '
' netUsePersistent - "yes" or "no" for the net use persistent state  '
'--------------------------------------------------------------------'

Function createSrvConfig(hostname, sharePaths, shareLetters, netUsePersistent)
	Set srvCfg = New ServerConfiguration
	srvCfg.hostname = hostname
	srvCfg.sharePaths = sharePaths
	srvCfg.shareLetters = shareLetters
	srvCfg.netUsePersistent = netUsePersistent
	Set createSrvConfig = srvCfg
End Function

'------------------------------------------------------------'
' Routine to initialize an array of server objects.          '
'                                                            '
' scriptConfig - object for the global script configuration  '
' srvConfigs - array of server configurations                '
'------------------------------------------------------------'

Sub initSrvs(ByVal scriptConfig, ByRef srvConfigs)
	For i = 0 to uBound(srvConfigs)
		initSrv scriptConfig, srvConfigs(i)
	Next
End Sub

'------------------------------------------------------------'
' Routine to initialize a server object.                     '
'                                                            '
' scriptConfig - object for the global script configuration  '
' srvConfig - configuration object of the server             '
'------------------------------------------------------------'

Sub initSrv(ByVal scriptConfig, ByRef srvConfig)
	setSrvPath srvConfig
	setSrvState scriptConfig, srvConfig
	srvConfig.connected = false
End Sub

'-----------------------------------------------------'
' Routine to set an absolute help server share path.  '
'                                                     '
' srvConfig - configuration object of the server      '
'-----------------------------------------------------'

Sub setSrvPath(ByRef srvConfig)
	srvConfig.hPath = "\\" & srvConfig.hostname & "\" & srvConfig.sharePaths(0)
End Sub

'--------------------------------------------------------------------'
' Routine to set server state by ping and share (FS) connectivity.   '
'                                                                    '
' scriptConfig - object for the global script configuration          '
' srvConfig - configuration object of the server                     '
'--------------------------------------------------------------------'

Sub setSrvState(ByVal scriptConfig, ByRef srvConfig)
	If Not retryPingServer(scriptConfig, srvConfig, true) Then
		srvConfig.online = scriptConfig.fso.FolderExists(srvConfig.hPath)
	Else
		srvConfig.online = false
	End If
End Sub

'-------------------------------------------------------------------'
' Routine that checks if there is an offline server in the array.   '
'                                                                   '
' srvConfigs - array of server configurations                       '
'-------------------------------------------------------------------'

Function isSrvOffline(srvConfigs)
	dim offline
	offline = false
	For i = 0 to uBound(srvConfigs)
		If Not srvConfigs(i).online Then
			offline = true
			Exit For
		End If
	Next
	isSrvOffline = offline
End Function

'------------------------------------------------------------------'
' Routine that checks if there is an online server in the array.   '
'                                                                  '
' srvConfigs - array of server configurations                      '
'------------------------------------------------------------------'

Function isSrvOnline(srvConfigs)
	dim online
	online = false
	For i = 0 to uBound(srvConfigs)
		If srvConfigs(i).online Then
			online = true
			Exit For
		End If
	Next
	isSrvOnline = online
End Function

'--------------------------------------------------------------'
' Routine that dynamically adjusts wait time on retry count.   '
'                                                              '
' scriptConfig - object for the global script configuration    '
' retries - number of already executed retries                 '
'--------------------------------------------------------------'

Function getReconWaitTime(scriptConfig, retries)
	Dim reconWait
	If retries <= 15 Then
		reconWait = scriptConfig.reconWait
	ElseIF retries > 15 And retries <= 30 Then
		reconWait = scriptConfig.reconWait * 4
	ElseIf retries > 30 And retries <= 45 Then
		reconWait = scriptConfig.reconWait * 6
	ElseIf retried > 45 And retries <= 60 Then
		reconWait = scriptConfig.reconWait * 8
	Else
		reconWait = scriptConfig.reconWait * 10
	End If
	getReconWaitTime = reconWait
End Function

'-------------------------------------------------------------------------'
' Routine that prints the debug output after every reconnect iteration.   '
'                                                                         '
' scriptConfig - object for the global script configuration               '
' onSrvs - output string for online servers                               '
' offSrvs - output string for offline servers                             '
'-------------------------------------------------------------------------'

Sub printDebug(scriptConfig, onSrvs, offSrvs)
	Dim debugOut
	If scriptConfig.debug Then
	debugOut = ""
		If Not (Len(onSrvs) = 0) Then
			debugOut = "Server(s) online:" & Mid(onSrvs, 3, Len(onSrvs)-1)
		End If
		If Not (Len(onSrvs) = 0) And Not (Len(offSrvs) = 0) Then
			debugOut = debugOut & vbNewLine
		End If
		If Not (Len(offSrvs) = 0) Then
			debugOut = debugOut & "Server(s) offline:" & Mid(offSrvs, 3, Len(offSrvs)-1)
		End If
		MsgBox(debugOut)
	End If
End Sub

'---------------------------------------------------------------'
' Routine to add a server identification to an output string.   '
'                                                               '
' scriptConfig - object for the global script configuration     '
' part - string on which the concat will be applied             '
'---------------------------------------------------------------'

Function getSrvDebugUnit(part, srvConfig)
	getSrvDebugUnit = part & " | " & srvConfig.hostname
End Function

'-------------------------------------------------------------------'
' Routine that checks server connectivity and tries to reconnect.   '
' shares if the server is online.                                   '
'                                                                   '
' scriptConfig - object for the global script configuration         '
' srvConfig - configuration object of the server                    '
'-------------------------------------------------------------------'

Sub checkNconnect(ByVal scriptConfig, ByRef srvConfig)
	If Not retryPingServer(scriptConfig, srvConfig, true) Then
		netUseServerShares scriptConfig, srvConfig
	End If
End Sub

'----------------------------------------------------------------------------'
' Routine to handle the connectivity of multiple servers and their shares.   '
'                                                                            '
' scriptConfig - object for the global script configuration                  '
' srvConfigs - array of server configuration objects                         '
'-----------------------------------------------------------------------------'

Sub waitOnServersConnect(scriptConfig, srvConfigs)
	Dim i, wait, offSrvs, onSrvs
		
	i = 0
	initConfig scriptConfig
	initSrvs scriptConfig, srvConfigs
	While ((i <= scriptConfig.serverRetryCtn - 1 And isSrvOffline(srvConfigs)) Or (i = 0 And isSrvOnline(srvConfigs)))
		offSrvs = ""
		onSrvs = ""
		wait = true
		i = i + 1
		
		For j = 0 to uBound(srvConfigs)
			If srvConfigs(j).online And Not srvConfigs(j).connected Then
				checkNconnect scriptConfig, srvConfigs(j)
				onSrvs = getSrvDebugUnit(onSrvs, srvConfigs(j))
			End If
		Next
		
		For j = 0 to uBound(srvConfigs)
			If Not srvConfigs(j).online Then
				If wait Then
					WScript.Sleep getReconWaitTime(scriptConfig, i)
					wait = false
				End If
				setSrvState scriptConfig, srvConfigs(j)
				If srvConfigs(j).online Then
					checkNconnect scriptConfig, srvConfigs(j)
					onSrvs = getSrvDebugUnit(onSrvs, srvConfigs(j))
				Else
					offSrvs = getSrvDebugUnit(offSrvs, srvConfigs(j))
				End If
			End If
		Next
		printDebug scriptConfig, onSrvs, offSrvs
	Wend
End Sub

'-----------------------------------------------------------------------------------'
' These are the OPTIONAL script parameters which can be adapted to TUNE             '
' the script if it reconnects to slow or to MINIMIZE the overhead.                  '
'                                                                                   '
' pingWait - wait time after failed server ping                                     '
' netUseWait - wait time after failed net use                                       '
' reconWait - wait time after failed availability check                             '
' pingCtn - how many pings per access request should be executed before giving up   '
' netUseCtn - how many net use fails per reconnect are allowed before giving up     '
' serverRetryCtn - how many overall reconnection tries should be executed           '
' pingTimeout - how many milliseconds pass before the ping is canceled              '
' debug - enable or disable debug messages on current reconnection state            '
'-----------------------------------------------------------------------------------'

Set scriptConfig = new ScriptConfiguration
scriptConfig.pingWait = 100
scriptConfig.netUseWait = 0
scriptConfig.reconWait = 2500
scriptConfig.pingCtn = 2
scriptConfig.netUseCtn = 1
scriptConfig.serverRetryCtn = 75
scriptConfig.pingTimeout = 200
scriptConfig.debug = false

'-------------------------------------------------------------------------------------'
' Here are the parameters which MUST be changed for your server(s) and remote         '
' share(s). If you want to use this script in your topology it's MANDATORY to         '
' change the values below APPROPRIATELY.                                              '
'                                                                                     '
' hostname - IP or hostname of the remote server (must be modified)                   '
' sharePaths - all share paths on the server (must be modified)                       '
' shareLetters - the share / drive letters for the defined paths (must be modified)   '
' netUsePersistent - should net use create a persistent share (yes/no)                '
'-------------------------------------------------------------------------------------'

'-------------------'
' Configure Servers '
'-------------------'

Dim srvConfigs

' Server Configuration explanation with parameters and variable names '
'set srvCfgX = createSrvConfig(hostname, sharePaths, shareLetters, netUsePersistent)

' Multi Server Configuration - two servers with two shares for each endpoint (please remove unnecessary lines) '
Set srvCfg1 = createSrvConfig("192.168.1.1", Array("path\to\share1", "path\to\share2"), Array("Z:", "Y:"), "yes")
Set srvCfg2 = createSrvConfig("192.168.1.2", Array("path\to\share3", "path\to\share3"), Array("X:", "W:"), "yes")

' add more server configurations here or remove them if needed (remove "srvCfg2" for single server configuration)
srvConfigs = Array(srvCfg1, srvCfg2) 

'--------------'
' Start Script '
'--------------'

waitOnServersConnect scriptConfig, srvConfigs
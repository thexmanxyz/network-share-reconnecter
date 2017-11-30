'--------------------------------------------------------------'
'   Network Share Reconnecter                                  '
'                                                              '
'   Purpose: This project tries to automatically reconnect     '
'            disconnected Windows network shares and drives    '
'            if they are offline or are listed as offline.     '
'            The current network and access state is           '
'            periodically checked until the server is          '
'            available or when the reconnection threshold      '
'            is hit without establishing any connectivity.     '
'                                                              '
'   Author: Andreas Kar (thex) <andreas.kar@gmx.at>            '
'--------------------------------------------------------------'

'-------------------------------------------------------------------------------------'
'                                                                                     '
'                                Configure Servers                                    '
'                                                                                     '
' Here are the parameters which MUST be changed for your server(s) and remote         '
' share(s). If you want to use this script in your topology it's MANDATORY to         '
' change the values below APPROPRIATELY.                                              '
'                                                                                     '
' hostname - IP, hostname or URI of the remote server (must be modified)              '
' sharePaths - all share paths on the server (must be modified)                       '
' shareLetters - the share / drive letters for the defined paths (must be modified)   '
' persistent - should net use create a persistent share (yes/no)                      '
' user - username for net use authentication if required (optional)                   '
' password - password for net use authentication if required (optional)               '
' secure - defines whether the HTTP or HTTPS protocol should be used (URI only)       '
'-------------------------------------------------------------------------------------'

Dim srvConfigs

'Server Configuration (UNC)'
'Set srvCfgX = createUncSrvConfig(hostname, sharePaths, shareLetters, persistent, username, password)

'Server Configuration (URI)'
'Set srvCfgY = createUriSrvConfig(hostname, sharePaths, shareLetters, persistent, username, password, secure)

'Multi Server Configuration - three servers with two shares for each endpoint [remove unnecessary lines]'
Set srvCfg1 = createUncSrvConfig("192.168.1.1", Array("path\to\share1", "path\to\share2"), Array("Z:", "Y:"), "yes", "", "")
Set srvCfg2 = createUncSrvConfig("192.168.1.2", Array("path\to\share3", "path\to\share4"), Array("X:", "W:"), "yes", "", "")
Set srvCfg3 = createUriSrvConfig("my.web.server", Array("path/to/share5", "path/to/share6"), Array("X:", "W:"), "yes", "", "", true)

'add more server configurations here or remove them if needed [remove srvCfg2 and srvCfg3 for single server configuration]'
srvConfigs = Array(srvCfg1, srvCfg2, srvCfg3) 

'-------------------------------------------------------------------------------------'
'                                                                                     '
'                                 Configure Script                                    '
'                                                                                     '
' These are the OPTIONAL script parameters which can be adapted to TUNE               '
' the script if it reconnects to slow or to MINIMIZE the overhead.                    '
'                                                                                     '
' pingWait - wait time after failed server ping                                       '
' netUseWait - wait time after failed net use                                         '
' reconWait - wait time after failed availability check                               '
' pingCtn - how many pings per access request should be executed before giving up     '
' netUseCtn - how many net use fails per reconnect are allowed before giving up       '
' serverRetryCtn - how many overall reconnection tries should be executed             '
' pingTimeout - how many milliseconds pass before the ping is canceled                '
' debug - enable or disable debug messages on current reconnection state              '
'-------------------------------------------------------------------------------------'

Set scriptConfig = new ScriptConfiguration
scriptConfig.pingWait = 100
scriptConfig.netUseWait = 0
scriptConfig.reconWait = 2500
scriptConfig.pingCtn = 2
scriptConfig.netUseCtn = 1
scriptConfig.serverRetryCtn = 75
scriptConfig.pingTimeout = 200
scriptConfig.debug = false

'--------------'
' Start Script '
'--------------'

waitOnServersConnect scriptConfig, srvConfigs

'--------------------------------------------'
' Simple class for the server configuration. '
'--------------------------------------------'

Class ServerConfiguration
	Public hostname
	Public sharePaths
	Public shareLetters
	Public persistent
	Public user
	Public password
	Public isUri
	Public secure
	Public online
	Public connected
	Public fsTestPath
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
			If srvConfigs(j).online Then
				If Not srvConfigs(j).connected Then
					checkNconnect scriptConfig, srvConfigs(j)
				End If
				If srvConfigs(j).connected Then
					onSrvs = getSrvDebug(onSrvs, srvConfigs(j))
				End If
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
					onSrvs = getSrvDebug(onSrvs, srvConfigs(j))
				Else
					offSrvs = getSrvDebug(offSrvs, srvConfigs(j))
				End If
			End If
		Next
		printDebug scriptConfig, onSrvs, offSrvs
	Wend
End Sub

'-------------------------------------------------------------------'
' Routine that checks server connectivity and tries to reconnect    '
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

'--------------------------------------------------------------------'
' Routine that creates a new UNC server object.                      '
'                                                                    '
' hostname - IP or human readable hostname                           '
' sharePaths - array that contains share paths                       '
' shareLetters - array that contains share letters                   '
' persistent - "yes" or "no" for the net use persistent state        '
' user - the username applied in the net use command                 '
' password - the password applied in the net use command             '
'--------------------------------------------------------------------'

Function createUncSrvConfig(hostname, sharePaths, shareLetters, persistent, user, password)
	Set createUncSrvConfig = createSrvConfig(hostname, sharePaths, shareLetters, persistent, user, password, false, false)
End Function

'--------------------------------------------------------------------'
' Routine that creates a new URI server object.                      '
'                                                                    '
' hostname - IP or URI                                               '
' sharePaths - array that contains share paths                       '
' shareLetters - array that contains share letters                   '
' persistent - "yes" or "no" for the net use persistent state        '
' user - the username applied in the net use command                 '
' password - the password applied in the net use command             '
' secure - true or false to use HTTP or HTTPS                        '
'--------------------------------------------------------------------'

Function createUriSrvConfig(hostname, sharePaths, shareLetters, persistent, user, password, secure)
	Set createUriSrvConfig = createSrvConfig(hostname, sharePaths, shareLetters, persistent, user, password, true, secure)
End Function

'--------------------------------------------------------------------'
' Routine that creates a new server object.                          '
'                                                                    '
' hostname - IP or URI                                               '
' sharePaths - array that contains share paths                       '
' shareLetters - array that contains share letters                   '
' persistent - "yes" or "no" for the net use persistent state        '
' isUri - is server accessible over UNC or URI                       '
' user - the username applied in the net use command                 '
' password - the password applied in the net use command             '
' secure - true or false to use HTTP or HTTPS                        '
'--------------------------------------------------------------------'

Function createSrvConfig(hostname, sharePaths, shareLetters, persistent, user, password, isUri, secure)
	Set srvCfg = New ServerConfiguration
	
	'trim share paths if necessary (remove leading and trailing slash)
	trimSharePaths sharePaths, isUri

	'create server configuration
	srvCfg.hostname = hostname
	srvCfg.sharePaths = sharePaths
	srvCfg.shareLetters = shareLetters
	srvCfg.persistent = persistent
	srvCfg.user = user
	srvCfg.password = password
	srvCfg.isUri = isUri
	srvCfg.secure = secure
	Set createSrvConfig = srvCfg
End Function

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
	Dim testPath, hPath
	hPath = trimSharePath(srvConfig.sharePaths(0), srvConfig.isUri)
	If(srvConfig.isUri) Then
		testPath = createUriPath(srvConfig.hostname, hPath, srvConfig.secure)
	Else
		testPath = createUncPath(srvConfig.hostname, hPath)
	End If
	srvConfig.fsTestPath = testPath
End Sub

'--------------------------------------------------------------------'
' Routine to create a UNC server address                             '
'                                                                    '
' host - the servers hostname or IP                                  '
' path - the path to the share                                       '
'--------------------------------------------------------------------'

Function createUncPath(host, path)
	createUncPath = "\\" & host & "\" & path
End Function

'--------------------------------------------------------------------'
' Routine to create a URI server address                             '
'                                                                    '
' host - the servers URI or IP                                       '
' path - the path to the share                                       '
' secure - true or false to use HTTP or HTTPS                        '
'--------------------------------------------------------------------'

Function createUriPath(host, path, secure)
	Dim protocol, hPath
	If secure Then
		protocol = "https"
	Else
		protocol = "http"
	End If
	
	If Len(path) > 0 Then
		hPath = "/" & path
	Else
		hPath = path
	End If
	
	createUriPath = protocol & "://" & host & hPath
End Function

'--------------------------------------------------------------------'
' Trims leading and Trailing slash from share paths if necessary.    '
'                                                                    '
' sharePaths - array that contains share paths                       '
' isUri - are UNC or URI paths passed                                '
'--------------------------------------------------------------------'

Function trimSharePaths(ByRef sharePaths, isUri)
	For j = 0 to uBound(sharePaths)
		sharePaths(j) = trimSharePath(sharePaths(j), isUri)
	Next
End Function

'--------------------------------------------------------------------'
' Trims leading and trailing slash from path if necessary.           '
'                                                                    '
' sharePath - single path to share                                   '
' isUri - is a URI or UNC path passed                                '
'--------------------------------------------------------------------'

Function trimSharePath(sharePath, isUri)
	Dim hPath, hLen, slash

	If isUri Then
		slash = "/"
	Else
		slash = "\"
	End If
	
	'remove leading slash
	hPath = sharePath
	hLen = Len(hPath)
	If hLen > 0 Then
		If InStr(1, hPath, slash) = 1 Then
			hPath = Mid(hPath, 2, hLen - 1)
		End If
	End If
	
	'remove trailing slash only if UNC
	hLen = Len(hPath)
	If Not isUri And hLen > 0 Then
		If InStr(hLen, hPath, slash) = hLen Then
			hPath = Mid(hPath, 1, hLen - 1)
		End If
	End If
	trimSharePath = hPath
End Function

'------------------------------------------------------------------'
' Routine to shell ping a server with a predefined configuration.  '
' (not used because ping does not return correct status)           '
'                                                                  '
' scriptConfig - object for the global script configuration        '
' srvConfig - configuration object of the server                   '
'------------------------------------------------------------------'

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
		online = Not isNull(pEle) And Not IsNull(pEle.StatusCode) And pEle.StatusCode = 0
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
' icmp - use icmp ping instead of shell ping                      '
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
	Dim address, user
	
	If srvConfig.isUri Then
		address = createUriPath(srvConfig.hostname, srvConfig.sharePaths(pos), srvConfig.secure)
	Else
		address = createUncPath(srvConfig.hostname, srvConfig.sharePaths(pos))
	End If
	
	If Len(srvConfig.user) > 0 Then
		user = " /user:" & srvConfig.user & " " & srvConfig.password
	else
		user = ""
	End If
	
	getNetUseCmd = "net use " & srvConfig.shareLetters(pos) & " " & chr(34) & address _ 
						& chr(34) & user & " /persistent:" & srvConfig.persistent
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
	srvConfig.connected = true
End Sub

'--------------------------------------------------------------------'
' Routine to set server state by ping and share (FS) connectivity.   '
'                                                                    '
' scriptConfig - object for the global script configuration          '
' srvConfig - configuration object of the server                     '
'--------------------------------------------------------------------'

Sub setSrvState(ByVal scriptConfig, ByRef srvConfig)
	If Not retryPingServer(scriptConfig, srvConfig, true) Then
		If srvConfig.isUri Then
			srvConfig.online = true
		Else
			srvConfig.online = scriptConfig.fso.FolderExists(srvConfig.fsTestPath)
		End If
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
	Dim offline
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
	Dim online
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
	Dim debugOut, onLen, offLen
	
	If scriptConfig.debug Then
		onLen = Len(onSrvs)
		offLen = Len(offSrvs)
		debugOut = ""
		
		If Not (onLen = 0) Then
			debugOut = "Server(s) online:" & Mid(onSrvs, 3, onLen - 1)
		End If
		If Not (onLen = 0) And Not (offLen = 0) Then
			debugOut = debugOut & vbNewLine
		End If
		If Not (offLen = 0) Then
			debugOut = debugOut & "Server(s) offline:" & Mid(offSrvs, 3, offLen - 1)
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

Function getSrvDebug(part, srvConfig)
	getSrvDebug = part & " | " & srvConfig.hostname
End Function
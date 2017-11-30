# Network Share Reconnecter
This project contains a **VB script** and a **Windows Scheduler Task** to automatically reconnect disconnected network shares / drives on login or unlock (if they are available or become available in the near future). But why? The reason for creating this project is that attached network drives are under Windows often not correctly reconnected on startup, restart or resume from hibernate / sleep. This still happens with the latest release *Windows 10 Fall Creators Update* or in my case especially with this release.

Many claim that the reconnect problem leads back to a timing problem during login and the establishment of the network connection over WLAN but there is no official statement from Microsoft as far as I know. There are many registry hacks out there which do not work and do not improve the described behaviour. Moreover this problem already exists for years so I now decided to create a reconnecter which reacts on network availability and share accessability.

![1](/screenshots/drives.png)

## Prerequisites
* Windows XP, 7, 8, 8.1, 10
* little knowledge on Windows Task Scheduling

## Download / Installation
1. [Download v1.3.0](https://github.com/thexmanxyz/network-share-reconnecter/releases/download/v1.3.0/nsr.v1.3.0.zip) of the Network Share Reconnecter Package.
2. Extract the files.
3. Modify the sample configuration in the **share_reconnect.vbs** script file.
   * At least modify `hostname`, `sharePaths` and `shareLetters` (see also the [configuration section](https://github.com/thexmanxyz/network-share-reconnecter/blob/master/src/share_reconnect.vbs#L41-L43) and [Issue #1](https://github.com/thexmanxyz/network-share-reconnecter/issues/1)).
   * Multiple servers can be easily added and configured.
4. Copy **share_reconnect.vbs** to a self defined directory.
5. Start Windows Task Scheduler - manually or with **taskschd.msc**.
6. Import **Network_Share_Reconnecter.xml**.
7. Modify the Scheduler Task.
   * At least change the path to the script which you have chosen before (or do it previously in the **Network_Share_Reconnecter.xml**).
   * (Optional) Change the user for the defined triggers, by default all computer users will be affected by the script.
   * (Optional) Extend or modify the Scheduler Task depending on your favor and preferences.

## Configuration and Parameters
Here a short description of the available parameters which can be configured:

* Server Configuration
  * `hostname` - IP or hostname of the remote server **(must be modified)**
  * `sharePaths` - all share paths on the server **(must be modified)**
  * `shareLetters` - the share / drive letters for the defined paths **(must be modified)**
  * `persistent` - should *net use* create a persistent share **(yes/no)**
  * `user` - username for net use authentication if required **(optional)**
  * `password` - password for net use authentication if required **(optional)**
  * `secure` - defines whether the HTTP or HTTPS protocol should be used **(URI only)**
  
* Global Script Configuration
  * `pingWait` - wait time after failed server ping
  * `netUseWait` - wait time after failed net use
  * `reconWait` - wait time after failed availability check
  * `pingCtn` - how many pings per access request should be executed before giving up
  * `netUseCtn` - how many *net use* fails per reconnect are allowed before giving up
  * `serverRetryCtn` - how many overall reconnection tries should be executed
  * `pingTimeout` - how many milliseconds pass before the ping is canceled
  * `debug` - enable or disable debug messages on current reconnection state

### UNC Example Configuration
If your share is accessible over an UNC path like `\\192.168.1.1\path\to\share` use this configuration.

`Set srvCfgUnc = createUncSrvConfig("192.168.1.1", Array("path\to\share"), Array("Z:"), "yes", "", "")`


### URI Example Configuration
If you share needs to be accessed over HTTP(s) like `http://my.webserver.com/path/to/share` use this configuration.

`Set srvCfgUri = createUncSrvConfig("my.webserver.com", Array("path/to/share"), Array("Z:"), "yes", "", "", true)`

## Features
* Automatic reconnection of network drives and shares on logon or unlock of the workstation.
* Stealth script execution (operates without exposing any prompts or windows - except when debug is enabled :P).
* Self determines server and share availability and accessability.
* Variable and flexible configuration of polling, timeouts and fallback handling.
* Configuration of multiple servers together with their shares.
* Fast ICMP ping checks and adaptive intensity.
* Scheduling task included (`.xml`) for easy import.

## Future Tasks
* net use analyze to better handle failure states
* install script or app
* permanent task / hook on event?

## Known Issues
None

## by [thex](https://github.com/thexmanxyz)
Copyright (c) 2017, free to use in personal and commercial software as per the [license](/LICENSE.md).

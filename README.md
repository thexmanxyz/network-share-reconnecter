# Network Share Reconnecter
This project contains a VB script and a Windows scheduler task to automatically reconnect disconnected network shares / drives on login (if they are available or become available). But why? The reason for creating this procject is that attached network drives are not correctly reconnected on startup, restart or resume from hibernate / sleep. This still happens in the latest release *Windows 10 Fall Creators Update* or in my case especially in this release.

Many claim that the reconnect problem leads back to a timing problem during login and the establishment of the network connection but there is no official statement from Microsoft as far as I know. There are many registry hacks out there which do not work and do not improve this behaviour. Moreover this problem already exists for years so I now decided to create a reconnecter which reacts on network availability and share accessability.

## Prerequisites
* Windows XP, 7, 8, 8.1, 10
* little knowledge on Windows Task Scheduling

## Download / Installation
1. [Download v1.0](https://github.com/thexmanxyz/network-share-reconnecter/releases/download/v1.0/nsr.v1.0.zip) of the Network Share Reconnecter Package
2. Extract the files
3. modify the **share_reconnect.vbs**
   * at least modify **hostname**, **sharePaths** and **shareLetters**
4. copy **share_reconnect.vbs** to a self defined directory
5. Start Windows Task Scheduler: manually or with **taskschd.msc**
6. Import **Network_Share_Reconnecter.xml**
7. Modify the Scheduler Task
   * change the path to the script you have chosen before (or do it previously in the **Network_Share_Reconnecter.xml**)
   * change the UserId for the defined Triggers (or do it previously in the **Network_Share_Reconnecter.xml**)
8. (Optional) change the Scheduler Task depending on your favor and preferences

If you need multi server support please wait until an improved version is out or otherwise creat multiple scheduler tasks and duplicate the script for each server. Otherwise you can also modify the script and call the *waitOnServerConnect()* routine multiple times (however this is the worst solution because it's not async and imposed a big delay).

## Configuration and Parameters
Here a short description of the available parameters which can be configured:

* Global Script Configuration
  * pingWait - wait time after failed server ping
  * netUseWait - wait time after failed net use
  * reconWait - wait time after failed availability check
  * pingCtn - how many pings per reconnect should be executed before giving up
  * netUseCtn - how many *net use* fails per reconnect are allowed before giving up
  * serverRetryCtn - how many overall reconnection tries should be executed
  * debug - enable or disable debug dialogs on current reconnection state

* Server Configuration
  * hostname - IP or hostname of the remote server **(must be modified)**
  * sharePaths - all share paths on the server **(must be modified)**
  * shareLetters - the share / drive letters for the defined paths **(must be modified)**
  * netUsePersistent - should *net use* create a persistent share **(yes/no)**

## Features
* automatic reconnection of network drives / shares on logon or unlock workstation
* stealth (operates without exposing any prompts or windows)
* determines availability and accessability 
* polling, retry and fallback
* full configuration of sleep and polling parameters
* scheduling task included (.xml) for easy import

## Future Tasks
* Better and more efficient multi-server support

## Known Issues
None

## by [thex](https://github.com/thexmanxyz)
Copyright (c) 2017, free to use in personal and commercial software as per the [license](/LICENSE.md).

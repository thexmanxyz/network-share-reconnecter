# Network Share Reconnecter
This project contains a VB script and a Windows scheduler task to automatically reconnect disconnected network shares / drives on login (if they are available or become available). But why? The reason for creating this procject is that attached network drives are not correctly reconnected on startup, restart or resume from hibernate / sleep. This still happens in the latest release *Windows 10 Fall Creators Update* or in my case especially in this release.

Many claim that the reconnect problem leads back to a timing problem during login and the establishment of the network connection but there is no official statement from Microsoft as fas as I know. There are many registry hacks out there which do not work and do not improve this behaviour. Moreover this problem already exists for years so I now decided to create a reconnecter which reacts on network availability and share accessability.

## Prerequisites
* Windows XP, 7, 8, 8.1, 10
* little knowledge on Windows Task Scheduling

## Download / Installation
1. [Download v1.0]() of the Network Share Reconnecter Package
2. Extract the files
3. modify the "share_reconnect.vbs"
   * at least modify "hostname", "sharePaths" and "shareLetters"
4. copy "share_reconnect.vbs" to a self defined directory
5. Start Windows Task Scheduler: manually or with "taskschd.msc"
6. Import "Network Share Reconnect.xml"
7. Modify the path to the script you have chosen before (or do it previously in the "Network Share Reconnect.xml")
8. (Optional) change the Scheduler Task depending on your favor and preferences

## Configuration and Parameters
Here a short description of the available parameters which can be configured:
* hostname - IP or hostname of server (must be modified)
* sharePaths - all share paths on that server (must be modified)
* shareLetters - the share / drive letters for the defined paths (must be modified)
* pingWait - wait time after failed server ping
* reconWait - wait time after failed availability check
* pingCtn - how many pings per reconnect should be executed before giving up
* netUseCtn - how many *net use* fails per reconnect are allowed before giving up
* serverRetryCtn - how many overall reconnection tries should be executed
* netUsePersistent - should *net use* create a persistent share (yes/no)
* debug - enable or disable debug dialogs on current reconnection state

## Features
* automatic reconnection of network drives / shares
* stealth (operates without exposing any prompts or windows)
* determines availability and accessability 
* polling, retry and fallback
* full configuration of sleep and polling parameters
* scheduling task included (.xml) for easy import

## Future Tasks
* asynchronous calls for different hosts
* Better multi-server support

## Known Issues
None

## by [thex](https://github.com/thexmanxyz)
Copyright (c) 2017, free to use in personal and commercial software as per the [license](/LICENSE.md).

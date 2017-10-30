# Network Share Reconnecter
This project contains a **VB script** and a **Windows Scheduler Task** to automatically reconnect disconnected network shares / drives on login or unlock (if they are available or become available in the near future). But why? The reason for creating this project is that attached network drives are not correctly reconnected on startup, restart or resume from hibernate / sleep. This still happens with the latest release *Windows 10 Fall Creators Update* or in my case especially with this release.

Many claim that the reconnect problem leads back to a timing problem during login and the establishment of the network connection with WLAN but there is no official statement from Microsoft as far as I know. There are many registry hacks out there which do not work and do not improve the described behaviour. Moreover this problem already exists for years so I now decided to create a reconnecter which reacts on network availability and share accessability.

![1](/screenshots/drives.png)

## Prerequisites
* Windows XP, 7, 8, 8.1, 10
* little knowledge on Windows Task Scheduling

## Download / Installation
1. [Download v1.2.0](https://github.com/thexmanxyz/network-share-reconnecter/releases/download/v1.2.0/nsr.v1.2.0.zip) of the Network Share Reconnecter Package
2. Extract the files
3. modify the **share_reconnect.vbs**
   * at least modify *hostname*, *sharePaths* and *shareLetters*
   * multiple servers can be easily added and configured
4. copy **share_reconnect.vbs** to a self defined directory
5. Start Windows Task Scheduler: manually or with **taskschd.msc**
6. Import **Network_Share_Reconnecter.xml**
7. Modify the Scheduler Task
   * at least change the path to the script, you have chosen before (or do it previously in the **Network_Share_Reconnecter.xml**)
   * (Optional) change the user for the defined triggers, by default all computer users will be affected by the script
8. (Optional) change the Scheduler Task depending on your favor and preferences

## Configuration and Parameters
Here a short description of the available parameters which can be configured:

* Global Script Configuration
  * pingWait - wait time after failed server ping
  * netUseWait - wait time after failed net use
  * reconWait - wait time after failed availability check
  * pingCtn - how many pings per access request should be executed before giving up
  * netUseCtn - how many *net use* fails per reconnect are allowed before giving up
  * serverRetryCtn - how many overall reconnection tries should be executed
  * pingTimeout - how many milliseconds pass before the ping is canceled
  * debug - enable or disable debug messages on current reconnection state

* Server Configuration
  * hostname - IP or hostname of the remote server **(must be modified)**
  * sharePaths - all share paths on the server **(must be modified)**
  * shareLetters - the share / drive letters for the defined paths **(must be modified)**
  * netUsePersistent - should *net use* create a persistent share **(yes/no)**

## Features
* automatic reconnection of network drives and shares on logon or unlock of the workstation
* stealth script execution (operates without exposing any prompts or windows - except when debug enabled :P)
* self determines server and share availability and accessability 
* variable and flexible configuration of polling, timeouts and fallback handling
* configuration of multiple servers together with their shares
* fast ICMP ping checks and adaptive intensity
* scheduling task included (.xml) for easy import

## Future Tasks
* net use analyze to better handle failure states

## Known Issues
None

## by [thex](https://github.com/thexmanxyz)
Copyright (c) 2017, free to use in personal and commercial software as per the [license](/LICENSE.md).

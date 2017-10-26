# Network Share Reconnecter
This project contains a VB script and a Windows scheduler task to automatically reconnect disconnected network shares / drives on login (if they are available or become available). But why? The reason for creating this procject is that *even in Windows 10 Fall Creators Update* attached network drives are not correctly reconnected on startup, restart or resume from hibernate / sleep. Many claim that the reconnect problem leads back to a timing problem during login and the establishment of the network connection but there is no official statement from Microsoft as fas as I know. There are many registry hacks out there which do not work and do not improve this behaviour. Moreover this problem already exists for years so I now decided to create a reconnecter which reacts on network availability and share accessability.

## Prerequisites
* Windows XP, 7, 8, 8.1, 10
* little knowledge on Windows Task Scheduling

## Download / Installation
1. a

## Configuration and Parameters

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

@echo off
REM ################################################################
REM #                                                              #
REM #   Network Share Reconnecter                                  #
REM #                                                              #
REM #   Purpose: This project tries to automatically reconnect     #
REM #            disconnected Windows network shares and drives    #
REM #            if they are offline or are listed as offline.     #
REM #            The current network and access state is           #
REM #            periodically checked until the server is          #
REM #            available or when the reconnection threshold      #
REM #            is hit without establishing any connectivity.     #
REM #                                                              #
REM #   Author: Andreas Kar (thex) <andreas.kar@gmx.at>            #
REM #                                                              #
REM ################################################################

REM --- Script Variables ---
set remove_folders=1
set log_files=0

REM --- Packaging Variables --
set prj_id=nsr
set prj_rev=v1.3.3
set prj_fullname=Network Share Reconnecter

REM --- File Variables ---
set prj_def_files=LICENSE, README.md

REM --- Folder Variables ---
set folder_root=..\..
set folder_src=src
set folder_task=task
set folder_temp=temp
set folder_release=release
set folder_def=default
set folder_releases=releases
set folder_release_dest=%folder_root%\%folder_releases%\%prj_rev%

REM --- Message Variables ---
set msg_start=Start build process for creating release.
set msg_finished=Successful finished build process.
set msg_success=successfully created.
set msg_release_success=Successful created Release "%prj_rev%". Packages moved to destination folder.
set msg_release_failed=Could not move packages to release folder "%prj_rev%", already exists.

REM --- Start Script ---
SETLOCAL EnableDelayedExpansion
for /F "tokens=1,2 delims=#" %%a in ('"prompt #$H#$E# & echo on & for %%b in (1) do rem"') do (
  set "DEL=%%a"
)

echo.
echo -----------------------------
echo # %prj_fullname% #
echo -----------------------------
echo.
echo %msg_start%

IF NOT EXIST %folder_temp% ( mkdir %folder_temp% )
IF NOT EXIST %folder_release% (	mkdir %folder_release% )

echo.
cd %folder_temp%

REM --- Call Package Creation ----
call :create_package "%folder_def%"

REM --- Move Packages to Release Folder ---
IF "%log_files%" == "0" ( echo. )
cd..
IF NOT EXIST %folder_release_dest% ( 
	mkdir %folder_release_dest%
	call :copy_folder_content "%folder_release%" "%folder_release_dest%"
	IF "%log_files%" == "1" ( echo ------------------------- )
	call :ColorizeText 0a "%msg_release_success%"
) ELSE (
	call :ColorizeText 0C "%msg_release_failed%"
)

REM --- Stop Script and Cleanup ---
IF %remove_folders% == 1 (
	rmdir "%folder_temp%" /S /Q
	rmdir "%folder_release%" /S /Q
)
call :ColorizeText 0a "%msg_finished%"
goto :EOF

REM --- Create Release Package ---
REM --- Parameters: %~1 = destination folder package
:create_package
	setlocal EnableDelayedExpansion	
		set folder_out=%~1
		set package_name=%prj_id%.%prj_rev%
		
		IF NOT EXIST !folder_out! ( mkdir !folder_out! )
		(for %%f in (%prj_def_files%) do ( call :copy_general_files "%folder_root%\%%f" "!folder_out!"	))
		call :copy_folder_content_ow "%folder_root%\%folder_src%" "!folder_out!"
		call :copy_folder_content_ow "%folder_root%\%folder_task%" "!folder_out!"
		call :create_archives "!package_name!" "!folder_out!" "1" "1"

		IF %remove_folders% == 1 ( rmdir "!folder_out!" /S /Q )
		IF "%log_files%" == "1" ( echo ------------------------- )
		echo !package_name! %msg_success%
		IF "%log_files%" == "1" ( echo. )
	endlocal
goto :EOF

REM --- Copies the general project files like license and readme  ---
REM --- Parameters: %~1 = source folder, %~2 = output folder
:copy_general_files
	IF "%log_files%" == "1" ( echo %~1 )
	copy %~1 %~2 >Nul
goto :EOF

REM --- Copy and Include a subfolder in the package ---
REM --- Parameters: %~1 = src folder path, %~2 = target folder path
:copy_include_sub_folder
	IF EXIST %folder_root%\%~1 (
		set folder_out_js=!folder_out!\%~2
		IF NOT EXIST !folder_out_js! ( mkdir !folder_out_js! )
		call :copy_folder_content_ow "%folder_root%\%~1" "!folder_out_js!"
	)
goto :EOF

REM --- Copies content of a folder and overwrites content
REM --- Parameters: %~1 = Source Folder, %~2 = Destination Folder
:copy_folder_content_ow
	IF "%log_files%" == "1" ( xcopy /s /Y %~1 %~2 ) ELSE ( xcopy /s /Y %~1 %~2 >Nul )
goto :EOF

REM --- Copies content of a folder without overwrite
REM --- Parameters: %~1 = Source Folder, %~2 = Destination Folder
:copy_folder_content
	IF "%log_files%" == "1" ( xcopy /s %~1 %~2 ) ELSE ( xcopy /s %~1 %~2 >Nul )
goto :EOF

REM --- Creates Release Archives ---
REM --- Parameters: %~1 = package name, %~2 = output folder, %~3 = create zip, %~4 = create tar.gz
:create_archives
	set arch_dest=..\%folder_release%\%~1
	
	IF "%~3" == "1" (
		set zip_dest=!arch_dest!.zip
		IF EXIST !zip_dest! ( del !zip_dest! )
		7z a -tzip !zip_dest! .\%~2\* >Nul
	)
	IF "%~4" == "1" (
		set tar_dest=!arch_dest!.tar
		set gzip_dest=!arch_dest!.tar.gz
		7z a -ttar !tar_dest! .\%~2\* >Nul
		IF EXIST !gzip_dest! ( del !gzip_dest! )
		7z a !gzip_dest! !tar_dest! >Nul
		del !tar_dest!
	)
goto :EOF

REM --- Colorizes Text ---
REM --- Parameters: %~1 = Color Hex, %~2 = Output
:ColorizeText
echo off
<nul set /p ".=%DEL%" > "%~2"
findstr /v /a:%1 /R "^$" "%~2" nul
del "%~2" > nul 2>&1
echo.
goto :EOF
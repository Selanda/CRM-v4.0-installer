@ECHO OFF
:: This batch file is designed to aid installation of the Microsoft CRM v4.0 Client

:MAIN_MENU
CLS
ECHO.
ECHO ********************************************
ECHO Please select an option from the list below:
ECHO ********************************************
ECHO.
ECHO ( 1 ) - Install Microsoft CRM v4.0 Client
ECHO ( 2 ) - Install CRM v4.0 Update Rollup 7
ECHO ( 3 ) - Install CRM v4.0 Update Rollup 8
ECHO ( 4 ) - Configuration Wizard (must be done under user's account)
ECHO ( 5 ) - Uninstall Microsoft CRM v4.0 Client
ECHO.
SET /P SELECTION=Make your selection now: 
ECHO.

IF %SELECTION%==1 GOTO OUTLOOK_STATUS_CHECK
IF %SELECTION%==2 GOTO INSTALL_UPDATE_7
IF %SELECTION%==3 GOTO INSTALL_UPDATE_8
IF %SELECTION%==4 GOTO LAUNCH_CONFIG
IF %SELECTION%==5 GOTO UNINSTALL_ROUTINE


:: Selection 1 - Install Microsoft CRM v4.0 Client
:: ------------------------------------------------
:OUTLOOK_STATUS_CHECK
ECHO Checking to see if Outlook is running...
TASKLIST /FI "IMAGENAME eq OUTLOOK.EXE" | FIND /I "outlook.exe" 
ECHO.
IF ERRORLEVEL 2 GOTO OUTLOOK_MURDER
IF ERRORLEVEL 1 GOTO INSTALL_CLIENT
	
:OUTLOOK_MURDER
TASKKILL /IM OUTLOOK.EXE
ECHO.
ECHO Outlook.exe has been successfully killed...
ECHO.
GOTO INSTALL_CLIENT
	
:INSTALL_CLIENT
ECHO Launching client install...
ECHO There are no status updates during this portion of the installation routine.
ECHO You will be returned to the main menu when the installation is complete.
ECHO.
ECHO Please wait...
ECHO.
ECHO Monitoring installation status...
I:\CRM\ClientInstall\SetupClient.exe /q
ECHO.

:SETUP_STATUS_CHECK
TASKLIST /FI "IMAGENAME eq CLIENTSETUP.EXE" | FIND /I "ClientSetup.exe"
IF ERRORLEVEL 2 GOTO SETUP_STATUS_CHECK
IF ERRORLEVEL 1 GOTO STARTMENU_STATUS_CHECK

:STARTMENU_STATUS_CHECK
IF EXIST "C:\Documents and Settings\All Users\Start Menu\Programs\Microsoft Dynamics CRM 4.0\Configuration Wizard.lnk" GOTO MAIN_MENU
GOTO STARTMENU_STATUS_CHECK


:: Selection 2 - Install CRM v4.0 Update Rollup 7 from I:\ drive
:: --------------------------------------------------------------
:INSTALL_UPDATE_7
ECHO.
ECHO Installing CRM v4.0 Update Rollup 7...
ECHO You will be returned to the main menu when the installation is complete.
ECHO Monitoring status of update executable...
"I:\CRM\CRM v4.0 Update Rollup 7 for Outlook Client\CRMv4.0-KB971782-i386-Client-ENU.exe" /q
ECHO.

:UPDATE7_STATUS_CHECK
TASKLIST /FI "IMAGENAME eq CRMv4.0-KB971782-i386-Client-ENU.exe" | FIND /I "CRMv4.0-KB971782-i386-Client-ENU.exe"
IF ERRORLEVEL 2 GOTO UPDATE7_STATUS_CHECK
IF ERRORLEVEL 1 GOTO MAIN_MENU


:: Selection 3 - Install CRM v4.0 Update Rollup 8 from I:\ drive
:: --------------------------------------------------------------
:INSTALL_UPDATE_8
ECHO.
ECHO Installing CRM v4.0 Update Rollup 8... 
ECHO You will be returned to the main menu when the installation is complete.
ECHO Monitoring status of update executable...
"I:\CRM\CRM v4.0 Update Rollup 8 for Outlook Client\CRMv4.0-KB975995-i386-Client-ENU.exe" /q
ECHO.

:UPDATE8_STATUS_CHECK
TASKLIST /FI "IMAGENAME eq CRMv4.0-KB975995-i386-Client-ENU.exe" | FIND /I "CRMv4.0-KB975995-i386-Client-ENU.exe"
IF ERRORLEVEL 2 GOTO UPDATE8_STATUS_CHECK
IF ERRORLEVEL 1 GOTO MAIN_MENU


:: Selection 4 - Run the Configuration Wizard
:: -------------------------------------------
:LAUNCH_CONFIG
:: Wait 5 seconds to ensure client installer is fully completed
PING -n 5 127.0.0.1 > NUL
ECHO.
ECHO Launching Configuration Wizard...
ECHO You will be returned to the main menu when the configuration wizard is finished.
ECHO.
CD "C:\Program Files\Microsoft Dynamics CRM\Client\ConfigWizard"
Microsoft.Crm.Client.Config.exe /config "C:\Documents and Settings\%username%\Desktop\CRM v4.0 Client\Client_Config.xml"
GOTO RESURRECT_OUTLOOK

:RESURRECT_OUTLOOK
:: Wait 10 seconds to ensure installer is finalized
PING -n 10 127.0.0.1 > NUL
ECHO Resurrecting Outlook to complete installation...
"C:\Program Files\Microsoft Office\Office12\OUTLOOK.EXE"
GOTO MAIN_MENU


:: Selection 5 - Re-run the installation utility and select 'Uninstall'
:: --------------------------------------------------------------------
:UNINSTALL_ROUTINE
ECHO Running uninstall utility...
I:\CRM\ClientInstall\SetupClient.exe /uninstall
:UNINSTALL_STATUS
TASKLIST /FI "IMAGENAME eq CLIENTSETUP.EXE" | FIND /I "ClientSetup.exe"
IF ERRORLEVEL 2 GOTO UNINSTALL_STATUS
IF ERRORLEVEL 1 GOTO MAIN_MENU
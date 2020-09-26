@ECHO OFF
:: This batch file is designed to aid installation of the Microsoft CRM v4.0 Client

:: Install Microsoft CRM v4.0 Client
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
ECHO.
ECHO Monitoring installation status...
I:\CRM\ClientInstall\SetupClient.exe /q
ECHO.
:SETUP_STATUS_CHECK
TASKLIST /FI "IMAGENAME eq CLIENTSETUP.EXE" | FIND /I "ClientSetup.exe"
IF ERRORLEVEL 2 GOTO SETUP_STATUS_CHECK
IF ERRORLEVEL 1 GOTO STARTMENU_STATUS_CHECK

:STARTMENU_STATUS_CHECK
IF EXIST "C:\Documents and Settings\All Users\Start Menu\Programs\Microsoft Dynamics CRM 4.0\Configuration Wizard.lnk" GOTO INSTALL_UPDATE_7
GOTO STARTMENU_STATUS_CHECK


:: Install CRM v4.0 Update Rollup 7 from I:\ drive
:: --------------------------------------------------------------
:INSTALL_UPDATE_7
:: Wait 10 seconds to ensure client installer is fully completed
PING -n 10 127.0.0.1 > NUL
ECHO.
ECHO Installing CRM v4.0 Update Rollup 7...
"I:\CRM\CRM v4.0 Update Rollup 7 for Outlook Client\CRMv4.0-KB971782-i386-Client-ENU.exe" /q
ECHO.
ECHO Monitoring status of update executable...
:UPDATE7_STATUS_CHECK
TASKLIST /FI "IMAGENAME eq CRMv4.0-KB971782-i386-Client-ENU.exe" | FIND /I "CRMv4.0-KB971782-i386-Client-ENU.exe"
IF ERRORLEVEL 2 GOTO UPDATE7_STATUS_CHECK
IF ERRORLEVEL 1 GOTO INSTALL_UPDATE_8


:: Install CRM v4.0 Update Rollup 8 from I:\ drive
:: --------------------------------------------------------------
:INSTALL_UPDATE_8
:: Wait 10 seconds to ensure client installer is fully completed
PING -n 10 127.0.0.1 > NUL
ECHO.
ECHO Installing CRM v4.0 Update Rollup 8... 
"I:\CRM\CRM v4.0 Update Rollup 8 for Outlook lient\CRMv4.0-KB975995-i386-Client-ENU.exe" /q
ECHO Monitoring status of update executable...
:UPDATE8_STATUS_CHECK
TASKLIST /FI "IMAGENAME eq CRMv4.0-KB975995-i386-Client-ENU.exe" | FIND /I "CRMv4.0-KB975995-i386-Client-ENU.exe"
IF ERRORLEVEL 2 GOTO UPDATE8_STATUS_CHECK
IF ERRORLEVEL 1 GOTO LAUNCH_CONFIG


:LAUNCH_CONFIG
ECHO.
ECHO Launching Configuration Wizard...
:: Wait 15 seconds to ensure update installer is fully completed
PING -n 15 127.0.0.1 > NUL
ECHO.
CD "C:\Program Files\Microsoft Dynamics CRM\Client\ConfigWizard"
Microsoft.Crm.Client.Config.exe /config "C:\Documents and Settings\%username%\Desktop\CRM v4.0 Client\Client_Config.xml"
:CONFIG_STATUS_CHECK
TASKLIST /FI "IMAGENAME eq Microsoft.Crm.Client.Config.exe" | FIND /I "Microsoft.Crm.Client.Config.exe"
IF ERRORLEVEL 2 GOTO CONFIG_STATUS_CHECK
IF ERRORLEVEL 1 GOTO RESURRECT_OUTLOOK

:RESURRECT_OUTLOOK
:: Wait 10 seconds to ensure configuration wizard is finalized
PING -n 10 127.0.0.1 > NUL
ECHO.
ECHO Setup of Microsoft CRM Client v4.0 for Outlook has completed.
ECHO.
ECHO Outlook will now be resurrected to complete the configuration...
"C:\Program Files\Microsoft Office\Office12\OUTLOOK.EXE"
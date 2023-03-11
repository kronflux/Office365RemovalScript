@CLS
@ECHO OFF

:init
setlocal DisableDelayedExpansion
set "batchPath=%~0"
FOR %%k IN (%0) DO set batchName=%%~nk
set "vbsGetPrivileges=%temp%\OEgetPriv_%batchName%.vbs"
setlocal EnableDelayedExpansion

:checkPrivileges
NET FILE 1>NUL 2>NUL
IF '%errorlevel%' == '0' ( GOTO gotPrivileges ) ELSE ( GOTO getPrivileges )

:getPrivileges
IF '%1'=='ELEV' (ECHO ELEV & shift /1 & GOTO gotPrivileges)
CLS
ECHO Invoking UAC for Privilege Elevation

ECHO Set UAC = CreateObject^("Shell.Application"^) > "%vbsGetPrivileges%"
ECHO args = "ELEV " >> "%vbsGetPrivileges%"
ECHO For Each strArg in WScript.Arguments >> "%vbsGetPrivileges%"
ECHO args = args ^& strArg ^& " "  >> "%vbsGetPrivileges%"
ECHO Next >> "%vbsGetPrivileges%"
ECHO UAC.ShellExecute "!batchPath!", args, "", "runas", 1 >> "%vbsGetPrivileges%"
"%SystemRoot%\System32\WScript.exe" "%vbsGetPrivileges%" %*
exit /B

:gotPrivileges
setlocal & pushd .
CD /D %~dp0
IF '%1'=='ELEV' (DEL "%vbsGetPrivileges%" 1>nul 2>nul  &  shift /1)

:--------------------------------------
setlocal enableextensions
echo.
echo ===============================================
echo Office 365 Removal Script
echo ===============================================
echo.
echo Press any key to run the script, or close
echo window to exit.
pause >nul 2>&1
echo Removing Microsoft Office 365...
echo.

:: Set process list for all known Office executables, including setup.exe and msiexec.exe
:: Note that this will also kill any running installers for other applications
set "processlist=accicons ai aimgr appsharinghookcontroller appvdllsurrogate appvdllsurrogate32 appvdllsurrogate64 appvlp appvshnotify bcssync clview cnfnot32"
set "processlist=%processlist% common.dbconnection common.dbconnection64 common.showhelp communicator databasecompare dbcicons dw20 eqnedt32 excel excelcnv filecompare"
set "processlist=%processlist% firstrun fltldr graph groove grv_icons iecontentservice iexplore infopath integratedoffice integrator joticon liclua lync lyncicon mavinject32"
set "processlist=%processlist% microsoft.mashup.container microsoft.mashup.container.loader microsoft.mashup.container.netfx40 microsoft.mashup.container.netfx45"
set "processlist=%processlist% misc msaccess msiexec msoadfsb msoasb msoev msohtmed msoia msoicons msosrec msosync msotd msoxmled mspub msqry32 namecontrolserver"
set "processlist=%processlist% officeappguardwin32 officec2rclient officeclicktorun officeondemand officesas officesasscheduler officescrbroker officescrsanbroker"
set "processlist=%processlist% ohub32 olcfg olicenseheartbeat onedrivesetup onenote onenotem operfmon orgchart ose osmadminicon osmclienticon ospprearm outicon outlook"
set "processlist=%processlist% pdfreflow perfboost pj11icon powerpnt pptico protocolhandler pubs roamingoffice scanpst sdxhelper sdxhelperbgt selfcert setlang setup"
set "processlist=%processlist% skypeserver smarttaginstall spreadsheetcompare sqldumper sscicons visicon vpreview werfault winword wordconv wordicon xlicons"

:: Kill all tasks in Tasklist that match the above Office process list
for /f "tokens=1" %%i in (
'TASKLIST ^|findstr /I /B "%processlist%"'
) Do ( 
    TASKKILL /T /IM "%%i" /F >nul 2>&1
)

net stop ose

:: Microsoft Office 15 & 16 Click-to-Run Components
start /wait msiexec /qn /norestart /x {20150000-008C-0000-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {20150000-008C-0C0A-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {50150000-008F-0000-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90150000-007E-0000-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90150000-007E-0000-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90150000-008C-0000-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90150000-008C-0000-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90150000-008C-0407-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90150000-008C-0409-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90150000-008C-040C-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90150000-008C-0413-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90150000-008C-0416-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90150000-008C-0804-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90150000-008C-0C0A-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-000F-0000-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-007E-0000-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-007E-0000-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0000-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0000-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0401-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0405-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0407-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0407-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0409-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0409-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-040B-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-040C-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-040C-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-040D-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0410-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0410-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0411-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0413-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0413-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0414-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0415-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0416-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0416-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0419-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0419-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-041A-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-041D-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-041F-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0804-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0816-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0816-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0C0A-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008C-0C0A-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-008F-0000-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {90160000-00DD-0000-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {95120000-00B9-0409-0000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {95120000-00B9-0409-1000-0000000FF1CE}
start /wait msiexec /qn /norestart /x {D535FC73-1F63-4347-896A-C97A45F11E9C}

:: Microsoft Office Activation Assistant
start /wait msiexec /qn /norestart /x {65DA2EC9-0642-47E9-AAE2-B5267AA14D75}
start /wait msiexec /qn /norestart /x {E50AE784-FABE-46DA-A1F8-7B6B56DCB22E}

sc query ose | findstr /I /C:"STOPPED"
If %errorlevel% EQU 0 (
    sc config ose start= disabled obj= LocalSystem
)

if exist "%PROGRAMFILES%\Microsoft Office 15" (
    takeown /f "%PROGRAMFILES%\Microsoft Office 15"
    attrib -r -s -h /s /d "%PROGRAMFILES%\Microsoft Office 15"
    rmdir /s /q "%PROGRAMFILES%\Microsoft Office 15"
) else if exist "%PROGRAMFILES(x86)%\Microsoft Office 15" (
    takeown /f "%PROGRAMFILES(x86)%\Microsoft Office 15"
    attrib -r -s -h /s /d "%PROGRAMFILES(x86)%\Microsoft Office 15"
    rmdir /s /q "%PROGRAMFILES(x86)%\Microsoft Office 15"
)

if exist "%PROGRAMFILES%\Microsoft Office 16" (
    takeown /f "%PROGRAMFILES%\Microsoft Office 16"
    attrib -r -s -h /s /d "%PROGRAMFILES%\Microsoft Office 16"
    rmdir /s /q "%PROGRAMFILES%\Microsoft Office 16"
) else if exist "%PROGRAMFILES(x86)%\Microsoft Office 16" (
    takeown /f "%PROGRAMFILES(x86)%\Microsoft Office 16"
    attrib -r -s -h /s /d "%PROGRAMFILES(x86)%\Microsoft Office 16"
    rmdir /s /q "%PROGRAMFILES(x86)%\Microsoft Office 16"
)

if exist "%PROGRAMFILES%\Microsoft Office\root\Integration" (
    cd /d "%PROGRAMFILES%\Microsoft Office\root\Integration"
    takeown /f C2RManifest*.xml
    attrib -r -s -h /s /d C2RManifest*.xml
    del /s /q /f C2RManifest*.xml
) else if exist "%PROGRAMFILES(x86)%\Microsoft Office\root\Integration" (
    cd /d "%PROGRAMFILES(x86)%\Microsoft Office\root\Integration"
    takeown /f C2RManifest*.xml
    attrib -r -s -h /s /d C2RManifest*.xml
    del /s /q /f C2RManifest*.xml
)
 
if exist "%PROGRAMFILES%\Microsoft Office\PackageManifest" (
    takeown /f "%PROGRAMFILES%\Microsoft Office\PackageManifest"
    attrib -r -s -h /s /d "%PROGRAMFILES%\Microsoft Office\PackageManifest"
    rmdir /s /q "%PROGRAMFILES%\Microsoft Office\PackageManifest"
) else if exist "%PROGRAMFILES(x86)%\Microsoft Office\PackageManifest" (
    takeown /f "%PROGRAMFILES(x86)%\Microsoft Office\PackageManifest"
    attrib -r -s -h /s /d "%PROGRAMFILES(x86)%\Microsoft Office\PackageManifest"
    rmdir /s /q "%PROGRAMFILES(x86)%\Microsoft Office\PackageManifest"
)

if exist "%PROGRAMFILES%\Microsoft Office\PackageSunrisePolicies" (
    takeown /f "%PROGRAMFILES%\Microsoft Office\PackageSunrisePolicies"
    attrib -r -s -h /s /d "%PROGRAMFILES%\Microsoft Office\PackageSunrisePolicies"
    rmdir /s /q "%PROGRAMFILES%\Microsoft Office\PackageSunrisePolicies"
) else if exist "%PROGRAMFILES(x86)%\Microsoft Office\PackageSunrisePolicies" (
    takeown /f "%PROGRAMFILES(x86)%\Microsoft Office\PackageSunrisePolicies"
    attrib -r -s -h /s /d "%PROGRAMFILES(x86)%\Microsoft Office\PackageSunrisePolicies"
    rmdir /s /q "%PROGRAMFILES(x86)%\Microsoft Office\PackageSunrisePolicies"
)

if exist "%PROGRAMFILES%\Microsoft Office\root" (
    takeown /f "%PROGRAMFILES%\Microsoft Office\root"
    attrib -r -s -h /s /d "%PROGRAMFILES%\Microsoft Office\root"
    rmdir /s /q "%PROGRAMFILES%\Microsoft Office\root"
) else if exist "%PROGRAMFILES(x86)%\Microsoft Office\root" (
    takeown /f "%PROGRAMFILES(x86)%\Microsoft Office\root"
    attrib -r -s -h /s /d "%PROGRAMFILES(x86)%\Microsoft Office\root"
    rmdir /s /q "%PROGRAMFILES(x86)%\Microsoft Office\root"
)

if exist "%PROGRAMFILES%\Microsoft Office\AppXManifest.xml" (
    takeown /f "%PROGRAMFILES%\Microsoft Office\AppXManifest.xml"
    attrib -r -s -h /s /d "%PROGRAMFILES%\Microsoft Office\AppXManifest.xml"
    del /s /q /f "%PROGRAMFILES%\Microsoft Office\AppXManifest.xml"
) else if exist "%PROGRAMFILES(x86)%\Microsoft Office\AppXManifest.xml" (
    takeown /f "%PROGRAMFILES(x86)%\Microsoft Office\AppXManifest.xml"
    attrib -r -s -h /s /d "%PROGRAMFILES(x86)%\Microsoft Office\AppXManifest.xml"
    del /s /q /f "%PROGRAMFILES(x86)%\Microsoft Office\AppXManifest.xml"
)

if exist "%PROGRAMFILES%\Microsoft Office\FileSystemMetadata.xml" (
    takeown /f "%PROGRAMFILES%\Microsoft Office\FileSystemMetadata.xml"
    attrib -r -s -h /s /d "%PROGRAMFILES%\Microsoft Office\FileSystemMetadata.xml"
    del /s /q /f "%PROGRAMFILES%\Microsoft Office\FileSystemMetadata.xml"
) else if exist "%PROGRAMFILES(x86)%\Microsoft Office\FileSystemMetadata.xml" (
    takeown /f "%PROGRAMFILES(x86)%\Microsoft Office\FileSystemMetadata.xml"
    attrib -r -s -h /s /d "%PROGRAMFILES(x86)%\Microsoft Office\FileSystemMetadata.xml"
    del /s /q /f "%PROGRAMFILES(x86)%\Microsoft Office\FileSystemMetadata.xml"
)

if exist "%PROGRAMDATA%\Microsoft\ClicToRun" (
    takeown /f "%PROGRAMDATA%\Microsoft\ClicToRun"
    attrib -r -s -h /s /d "%PROGRAMDATA%\Microsoft\ClicToRun"
    rmdir /s /q "%PROGRAMDATA%\Microsoft\ClicToRun"
)

if exist "%COMMONPROGRAMFILES%\Microsoft Shared\ClickToRun" (
    takeown /f "%COMMONPROGRAMFILES%\Microsoft Shared\ClickToRun"
    attrib -r -s -h /s /d "%COMMONPROGRAMFILES%\Microsoft Shared\ClickToRun"
    rmdir /s /q "%COMMONPROGRAMFILES%\Microsoft Shared\ClickToRun"
) else if exist "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\ClickToRun" (
    takeown /f "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\ClickToRun"
    attrib -r -s -h /s /d "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\ClickToRun"
    rmdir /s /q "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\ClickToRun"
)

if exist "%PROGRAMDATA%\Microsoft\office\FFPackageLocker" (
    takeown /f "%PROGRAMDATA%\Microsoft\office\FFPackageLocker"
    attrib -r -s -h /s /d "%PROGRAMDATA%\Microsoft\office\FFPackageLocker"
    rmdir /s /q "%PROGRAMDATA%\Microsoft\office\FFPackageLocker"
)

if exist "%PROGRAMDATA%\Microsoft\office\ClickToRunPackageLocker" (
    takeown /f "%PROGRAMDATA%\Microsoft\office\ClickToRunPackageLocker"
    attrib -r -s -h /s /d "%PROGRAMDATA%\Microsoft\office\ClickToRunPackageLocker"
    rmdir /s /q "%PROGRAMDATA%\Microsoft\office\ClickToRunPackageLocker"
)

if exist "%PROGRAMDATA%\Microsoft\office\FFStatePBLocker" (
    takeown /f "%PROGRAMDATA%\Microsoft\office\FFStatePBLocker"
    attrib -r -s -h /s /d "%PROGRAMDATA%\Microsoft\office\FFStatePBLocker"
    rmdir /s /q "%PROGRAMDATA%\Microsoft\office\FFStatePBLocker"
)

if exist "%PROGRAMDATA%\Microsoft\office\Heartbeat" (
    takeown /f "%PROGRAMDATA%\Microsoft\office\Heartbeat"
    attrib -r -s -h /s /d "%PROGRAMDATA%\Microsoft\office\Heartbeat"
    rmdir /s /q "%PROGRAMDATA%\Microsoft\office\Heartbeat"
)

if exist "%USERPROFILE%\Microsoft Office" (
    takeown /f "%USERPROFILE%\Microsoft Office"
    attrib -r -s -h /s /d "%USERPROFILE%\Microsoft Office"
    rmdir /s /q "%USERPROFILE%\Microsoft Office"
)

if exist "%USERPROFILE%\Microsoft Office 15" (
    takeown /f "%USERPROFILE%\Microsoft Office 15"
    attrib -r -s -h /s /d "%USERPROFILE%\Microsoft Office 15"
    rmdir /s /q "%USERPROFILE%\Microsoft Office 15"
)

if exist "%USERPROFILE%\Microsoft Office 16" (
    takeown /f "%USERPROFILE%\Microsoft Office 16"
    attrib -r -s -h /s /d "%USERPROFILE%\Microsoft Office 16"
    rmdir /s /q "%USERPROFILE%\Microsoft Office 16"
)

schtasks.exe /delete /tn "FF_INTEGRATEDstreamSchedule" /f
schtasks.exe /delete /tn "FF_INTEGRATEDUPDATEDETECTION" /f
schtasks.exe /delete /tn "C2RAppVLoggingStart" /f
schtasks.exe /delete /tn "Office 15 Subscription Heartbeat" /f
schtasks.exe /delete /tn "\Microsoft\Office\Office 15 Subscription Heartbeat" /f
schtasks.exe /delete /tn "Microsoft Office 15 Sync Maintenance for {d068b555-9700-40b8-992c-f866287b06c1}" /f
schtasks.exe /delete /tn "\Microsoft\Office\OfficeInventoryAgentFallBack" /f
schtasks.exe /delete /tn "\Microsoft\Office\OfficeTelemetryAgentFallBack" /f
schtasks.exe /delete /tn "\Microsoft\Office\OfficeInventoryAgentLogOn" /f
schtasks.exe /delete /tn "\Microsoft\Office\OfficeTelemetryAgentLogOn" /f
schtasks.exe /delete /tn "Office Background Streaming" /f
schtasks.exe /delete /tn "\Microsoft\Office\Office Automatic Updates" /f
schtasks.exe /delete /tn "\Microsoft\Office\Office ClickToRun Service Monitor" /f
schtasks.exe /delete /tn "Office Subscription Maintenance" /f
schtasks.exe /delete /tn "\Microsoft\Office\Office Subscription Maintenance" /f
schtasks.exe /delete /tn "\Microsoft\Office\OfficeTelemetryAgentLogOn2016"
schtasks.exe /delete /tn "\Microsoft\Office\OfficeTelemetryAgentFallBack2016"

sc delete Clicktorunsvc

reg delete "HKCU\SOFTWARE\Microsoft\Office\15.0\ClickToRun" /f
reg delete "HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun" /f

reg delete "HKCU\SOFTWARE\Microsoft\Office\15.0\ClickToRunStore" /f
reg delete "HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRunStore" /f

reg delete "HKCU\SOFTWARE\Microsoft\Office\16.0\ClickToRun" /f
reg delete "HKLM\SOFTWARE\Microsoft\Office\16.0\ClickToRun" /f

reg delete "HKCU\SOFTWARE\Microsoft\Office\16.0\ClickToRunStore" /f
reg delete "HKLM\SOFTWARE\Microsoft\Office\16.0\ClickToRunStore" /f

reg delete "HKCU\SOFTWARE\Microsoft\Office\ClickToRun" /f
reg delete "HKLM\SOFTWARE\Microsoft\Office\ClickToRun" /f

reg delete "HKCU\SOFTWARE\Microsoft\Office\ClickToRunStore" /f
reg delete "HKLM\SOFTWARE\Microsoft\Office\ClickToRunStore" /f

reg delete "HKCU\Software\Microsoft\Office\15.0\Registration" /f
reg delete "HKLM\Software\Microsoft\Office\15.0\Registration" /f

reg delete "HKCU\Software\Microsoft\Office\16.0\Registration" /f
reg delete "HKLM\Software\Microsoft\Office\16.0\Registration" /f

reg delete "HKCU\Software\Microsoft\Office\Registration" /f
reg delete "HKLM\Software\Microsoft\Office\Registration" /f

reg delete "HKCU\Software\Policies\Microsoft\Office" /f
reg delete "HKLM\Software\Policies\Microsoft\Office" /f

reg delete "HKCU\Software\Microsoft\Office\15.0" /f
reg delete "HKLM\Software\Microsoft\Office\15.0" /f

reg delete "HKCU\Software\Microsoft\Office\16.0" /f
reg delete "HKLM\Software\Microsoft\Office\16.0" /f

reg delete "HKCU\Software\Microsoft\Office\Excel" /f
reg delete "HKLM\Software\Microsoft\Office\Excel" /f

reg delete "HKCU\Software\Microsoft\Office\Outlook" /f
reg delete "HKLM\Software\Microsoft\Office\Outlook" /f

reg delete "HKCU\Software\Microsoft\Office\Powerpoint" /f
reg delete "HKLM\Software\Microsoft\Office\Powerpoint" /f

reg delete "HKCU\Software\Microsoft\Office\Word" /f
reg delete "HKLM\Software\Microsoft\Office\Word" /f

reg delete "HKCU\Software\Microsoft\Office\Common" /f
reg delete "HKLM\Software\Microsoft\Office\Common" /f

reg delete "HKCU\SOFTWARE\Microsoft\Office" /f
reg delete "HKLM\SOFTWARE\Microsoft\Office" /f

reg delete "HKLM\SOFTWARE\Classes\CLSID\{2027FC3B-CF9D-4ec7-A823-38BA308625CC}" /f
reg delete "HKLM\SOFTWARE\Classes\Protocols\Handler\osf" /f
reg delete "HKLM\SOFTWARE\Microsoft\AppVISV" /f
reg delete "HKLM\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot\Virtual" /f
reg delete "HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot\Virtual" /f
reg delete "HKLM\SOFTWARE\Microsoft\Office\Common\InstallRoot\Virtual" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{31D09BA0-12F5-4CCE-BE8A-2923E76605DA}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{B4F3A835-0E21-4959-BA22-42B3008E02FF}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{0875DCB6-C686-4243-9432-ADCCF0B9F2D7}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\Namespace\{B28AA736-876B-46DA-B3A8-84C5E30BA492}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\NetworkNeighborhood\Namespace\{46137B78-0EC3-426D-8B89-FF7C3A458B5E}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 1 (ErrorConflict)" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 2 (SyncInProgress)" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 3 (InSync)" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Microsoft Office Temp Files" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\Lync15" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\Lync16" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\{0875DCB6-C686-4243-9432-ADCCF0B9F2D7}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\{42042206-2D85-11D3-8CFF-005004838597}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\{46137B78-0EC3-426D-8B89-FF7C3A458B5E}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\{506F4668-F13E-4AA1-BB04-B43203AB3CC0}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\{8B02D659-EBBB-43D7-9BBA-52CF22C5B025}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\{993BE281-6695-4BA5-8A2A-7AACBFAAB69E}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\{B28AA736-876B-46DA-B3A8-84C5E30BA492}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\{C41662BB-1FA0-4CE0-8DC5-9B7F8279FF97}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\{D66DC78C-4F61-447F-942B-3FB6980118CF}" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{31D09BA0-12F5-4CCE-BE8A-2923E76605DA}" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{B4F3A835-0E21-4959-BA22-42B3008E02FF}" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 1 (ErrorConflict)" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 2 (SyncInProgress)" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 3 (InSync)" /f

set clave=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
set valor=0FF1CE

for /f %%a in ('"reg query "%clave%" | find "%valor%""') do (
    reg delete "%%a" /f
)

set clave=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
set valor=O365

for /f %%a in ('"reg query "%clave%" | find "%valor%""') do (
    reg delete "%%a - en-us" /f
)

set clave=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
set valor=ProfessionaRetail

for /f %%a in ('"reg query "%clave%" | find "%valor%""') do (
    reg delete "%%a - en-us" /f
)

set clave=HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall
set valor=0FF1CE

for /f %%a in ('"reg query "%clave%" | find "%valor%""') do (
    reg delete "%%a" /f
)

set clave=HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall
set valor=O365

for /f %%a in ('"reg query "%clave%" | find "%valor%""') do (
    reg delete "%%a - en-us" /f
)

set clave=HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall
set valor=ProfessionalRetail

for /f %%a in ('"reg query "%clave%" | find "%valor%""') do (
    reg delete "%%a - en-us" /f
)

if exist "%PROGRAMFILES%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016" (
    takeown /f "%PROGRAMFILES%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016"
    attrib -r -s -h /s /d "%PROGRAMFILES%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016"
    rmdir /s /q "%PROGRAMFILES%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016"
)

if exist "%PROGRAMFILES%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016 Tools" (
    takeown /f "%PROGRAMFILES%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016 Tools"
    attrib -r -s -h /s /d "%PROGRAMFILES%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016 Tools"
    rmdir /s /q "%PROGRAMFILES%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016 Tools"
)

cd /d "%PROGRAMFILES%\Microsoft\Windows\Start Menu\Programs"
takeown /f *2016.lnk
attrib -r -s -h /s /d *2016.lnk
del /s /f /q *2016.lnk
 
if exist "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016" (
    takeown /f "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016"
    attrib -r -s -h /s /d "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016"
    rmdir /s /q "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016"
)

if exist "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016 Tools" (
    takeown /f "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016 Tools"
    attrib -r -s -h /s /d "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016 Tools"
    rmdir /s /q "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016 Tools"
)

cd /d "%APPDATA%\Microsoft\Windows\Start Menu\Programs"
takeown /f *2016.lnk
attrib -r -s -h /s /d *2016.lnk
del /s /f /q *2016.lnk
 
if exist "%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016" (
    takeown /f "%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016"
    attrib -r -s -h /s /d "%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016"
    rmdir /s /q "%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016"
)

if exist "%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016 Tools" (
    takeown /f "%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016 Tools"
    attrib -r -s -h /s /d "%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016 Tools"
    rmdir /s /q "%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016 Tools"
)

cd /d "%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs"
takeown /f *2016.lnk
attrib -r -s -h /s /d *2016.lnk
del /s /f /q *2016.lnk
 
if exist "%USERPROFILE%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016" (
    takeown /f "%USERPROFILE%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016"
    attrib -r -s -h /s /d "%USERPROFILE%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016"
    rmdir /s /q "%USERPROFILE%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016"
)

if exist "%USERPROFILE%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016 Tools" (
    takeown /f "%USERPROFILE%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016 Tools"
    attrib -r -s -h /s /d "%USERPROFILE%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016 Tools"
    rmdir /s /q "%USERPROFILE%\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016 Tools"
)

cd /d "%USERPROFILE%\Microsoft\Windows\Start Menu\Programs"
takeown /f *2016.lnk
attrib -r -s -h /s /d *2016.lnk
del /s /f /q *2016.lnk

if exist "%PROGRAMFILES%\Microsoft Office 15" (
    takeown /f "%PROGRAMFILES%\Microsoft Office 15"
    attrib -r -s -h /s /d "%PROGRAMFILES%\Microsoft Office 15"
    rmdir /s /q "%PROGRAMFILES%\Microsoft Office 15"
) else if exist "%PROGRAMFILES(x86)%\Microsoft Office 15" (
    takeown /f "%PROGRAMFILES(x86)%\Microsoft Office 15"
    attrib -r -s -h /s /d "%PROGRAMFILES(x86)%\Microsoft Office 15"
    rmdir /s /q "%PROGRAMFILES(x86)%\Microsoft Office 15"
)

if exist "%PROGRAMFILES%\Microsoft Office\root\Integration" (
    cd /d "%PROGRAMFILES%\Microsoft Office\root\Integration"
    takeown /f C2RManifest*.xml
    attrib -r -s -h /s /d C2RManifest*.xml
    del /s /q /f C2RManifest*.xml
) else if exist "%PROGRAMFILES(x86)%\Microsoft Office\root\Integration" (
    cd /d "%PROGRAMFILES(x86)%\Microsoft Office\root\Integration"
    takeown /f C2RManifest*.xml
    attrib -r -s -h /s /d C2RManifest*.xml
    del /s /q /f C2RManifest*.xml
)
 
if exist "%PROGRAMFILES%\Microsoft Office\PackageManifest" (
    takeown /f "%PROGRAMFILES%\Microsoft Office\PackageManifest"
    attrib -r -s -h /s /d "%PROGRAMFILES%\Microsoft Office\PackageManifest"
    rmdir /s /q "%PROGRAMFILES%\Microsoft Office\PackageManifest"
) else if exist "%PROGRAMFILES(x86)%\Microsoft Office\PackageManifest" (
    takeown /f "%PROGRAMFILES(x86)%\Microsoft Office\PackageManifest"
    attrib -r -s -h /s /d "%PROGRAMFILES(x86)%\Microsoft Office\PackageManifest"
    rmdir /s /q "%PROGRAMFILES(x86)%\Microsoft Office\PackageManifest"
)

if exist "%PROGRAMFILES%\Microsoft Office\PackageSunrisePolicies" (
    takeown /f "%PROGRAMFILES%\Microsoft Office\PackageSunrisePolicies"
    attrib -r -s -h /s /d "%PROGRAMFILES%\Microsoft Office\PackageSunrisePolicies"
    rmdir /s /q "%PROGRAMFILES%\Microsoft Office\PackageSunrisePolicies"
) else if exist "%PROGRAMFILES(x86)%\Microsoft Office\PackageSunrisePolicies" (
    takeown /f "%PROGRAMFILES(x86)%\Microsoft Office\PackageSunrisePolicies"
    attrib -r -s -h /s /d "%PROGRAMFILES(x86)%\Microsoft Office\PackageSunrisePolicies"
    rmdir /s /q "%PROGRAMFILES(x86)%\Microsoft Office\PackageSunrisePolicies"
)

if exist "%PROGRAMFILES%\Microsoft Office\root" (
    takeown /f "%PROGRAMFILES%\Microsoft Office\root"
    attrib -r -s -h /s /d "%PROGRAMFILES%\Microsoft Office\root"
    rmdir /s /q "%PROGRAMFILES%\Microsoft Office\root"
) else if exist "%PROGRAMFILES(x86)%\Microsoft Office\root" (
    takeown /f "%PROGRAMFILES(x86)%\Microsoft Office\root"
    attrib -r -s -h /s /d "%PROGRAMFILES(x86)%\Microsoft Office\root"
    rmdir /s /q "%PROGRAMFILES(x86)%\Microsoft Office\root"
)

if exist "%PROGRAMFILES%\Microsoft Office\AppXManifest.xml" (
    takeown /f "%PROGRAMFILES%\Microsoft Office\AppXManifest.xml"
    attrib -r -s -h /s /d "%PROGRAMFILES%\Microsoft Office\AppXManifest.xml"
    del /s /q /f "%PROGRAMFILES%\Microsoft Office\AppXManifest.xml"
) else if exist "%PROGRAMFILES(x86)%\Microsoft Office\AppXManifest.xml" (
    takeown /f "%PROGRAMFILES(x86)%\Microsoft Office\AppXManifest.xml"
    attrib -r -s -h /s /d "%PROGRAMFILES(x86)%\Microsoft Office\AppXManifest.xml"
    del /s /q /f "%PROGRAMFILES(x86)%\Microsoft Office\AppXManifest.xml"
)

if exist "%PROGRAMFILES%\Microsoft Office\FileSystemMetadata.xml" (
    takeown /f "%PROGRAMFILES%\Microsoft Office\FileSystemMetadata.xml"
    attrib -r -s -h /s /d "%PROGRAMFILES%\Microsoft Office\FileSystemMetadata.xml"
    del /s /q /f "%PROGRAMFILES%\Microsoft Office\FileSystemMetadata.xml"
) else if exist "%PROGRAMFILES(x86)%\Microsoft Office\FileSystemMetadata.xml" (
   takeown /f "%PROGRAMFILES(x86)%\Microsoft Office\FileSystemMetadata.xml"
   attrib -r -s -h /s /d "%PROGRAMFILES(x86)%\Microsoft Office\FileSystemMetadata.xml"
   del /s /q /f "%PROGRAMFILES(x86)%\Microsoft Office\FileSystemMetadata.xml"
)

if exist "%PROGRAMDATA%\Microsoft\ClicToRun" (
    takeown /f "%PROGRAMDATA%\Microsoft\ClicToRun"
    attrib -r -s -h /s /d "%PROGRAMDATA%\Microsoft\ClicToRun"
    rmdir /s /q "%PROGRAMDATA%\Microsoft\ClicToRun"
)

if exist "%COMMONPROGRAMFILES%\Microsoft Shared\ClickToRun" (
    takeown /f "%COMMONPROGRAMFILES%\Microsoft Shared\ClickToRun"
    attrib -r -s -h /s /d "%COMMONPROGRAMFILES%\Microsoft Shared\ClickToRun"
    rmdir /s /q "%COMMONPROGRAMFILES%\Microsoft Shared\ClickToRun"
) else if exist "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\ClickToRun" (
    takeown /f "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\ClickToRun"
    attrib -r -s -h /s /d "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\ClickToRun"
    rmdir /s /q "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\ClickToRun"
)

if exist "%PROGRAMDATA%\Microsoft\office\FFPackageLocker" (
    takeown /f "%PROGRAMDATA%\Microsoft\office\FFPackageLocker"
    attrib -r -s -h /s /d "%PROGRAMDATA%\Microsoft\office\FFPackageLocker"
    rmdir /s /q "%PROGRAMDATA%\Microsoft\office\FFPackageLocker"
)

if exist "%PROGRAMDATA%\Microsoft\office\ClickToRunPackageLocker" (
    takeown /f "%PROGRAMDATA%\Microsoft\office\ClickToRunPackageLocker"
    attrib -r -s -h /s /d "%PROGRAMDATA%\Microsoft\office\ClickToRunPackageLocker"
    rmdir /s /q "%PROGRAMDATA%\Microsoft\office\ClickToRunPackageLocker"
)

if exist "%PROGRAMDATA%\Microsoft\office\FFStatePBLocker" (
    takeown /f "%PROGRAMDATA%\Microsoft\office\FFStatePBLocker"
    attrib -r -s -h /s /d "%PROGRAMDATA%\Microsoft\office\FFStatePBLocker"
    rmdir /s /q "%PROGRAMDATA%\Microsoft\office\FFStatePBLocker"
)

if exist "%PROGRAMDATA%\Microsoft\office\Heartbeat" (
    takeown /f "%PROGRAMDATA%\Microsoft\office\Heartbeat"
    attrib -r -s -h /s /d "%PROGRAMDATA%\Microsoft\office\Heartbeat"
    rmdir /s /q "%PROGRAMDATA%\Microsoft\office\Heartbeat"
)

if exist "%USERPROFILE%\Microsoft Office" (
    takeown /f "%USERPROFILE%\Microsoft Office"
    attrib -r -s -h /s /d "%USERPROFILE%\Microsoft Office"
    rmdir /s /q "%USERPROFILE%\Microsoft Office"
)

if exist "%USERPROFILE%\Microsoft Office 15" (
    takeown /f "%USERPROFILE%\Microsoft Office 15"
    attrib -r -s -h /s /d "%USERPROFILE%\Microsoft Office 15"
    rmdir /s /q "%USERPROFILE%\Microsoft Office 15"
)

if exist "%USERPROFILE%\Microsoft Office 16" (
    takeown /f "%USERPROFILE%\Microsoft Office 16"
    attrib -r -s -h /s /d "%USERPROFILE%\Microsoft Office 16"
    rmdir /s /q "%USERPROFILE%\Microsoft Office 16"
)

if exist "%PROGRAMFILES%\Microsoft Office" (
    takeown /f "%PROGRAMFILES%\Microsoft Office"
    attrib -r -s -h /s /d "%PROGRAMFILES%\Microsoft Office"
    rmdir /s /q "%PROGRAMFILES%\Microsoft Office"
) else if exist "%PROGRAMFILES(x86)%\Microsoft Office" (
    takeown /f "%PROGRAMFILES(x86)%\Microsoft Office"
    attrib -r -s -h /s /d "%PROGRAMFILES(x86)%\Microsoft Office"
    rmdir /s /q "%PROGRAMFILES(x86)%\Microsoft Office"
)

if exist "%COMMONPROGRAMFILES%\Microsoft Shared\Office11" (
    takeown /f "%COMMONPROGRAMFILES%\Microsoft Shared\Office11"
    attrib -r -s -h /s /d "%COMMONPROGRAMFILES%\Microsoft Shared\Office11"
    rmdir /s /q "%COMMONPROGRAMFILES%\Microsoft Shared\Office11"
) else if exist "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\Office11" (
    takeown /f "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\Office11"    
	attrib -r -s -h /s /d "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\Office11"
    rmdir /s /q "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\Office11"
)

if exist "%COMMONPROGRAMFILES%\Microsoft Shared\Office12" (
    takeown /f "%COMMONPROGRAMFILES%\Microsoft Shared\Office12"
    attrib -r -s -h /s /d "%COMMONPROGRAMFILES%\Microsoft Shared\Office12"
    rmdir /s /q "%COMMONPROGRAMFILES%\Microsoft Shared\Office12"
) else if exist "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\Office12" (
    takeown /f "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\Office12"
    attrib -r -s -h /s /d "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\Office12"
    rmdir /s /q "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\Office12"
)

if exist "%COMMONPROGRAMFILES%\Microsoft Shared\Office14" (
    takeown /f "%COMMONPROGRAMFILES%\Microsoft Shared\Office14"
    attrib -r -s -h /s /d "%COMMONPROGRAMFILES%\Microsoft Shared\Office14"
    rmdir /s /q "%COMMONPROGRAMFILES%\Microsoft Shared\Office14"
) else if exist "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\Office14" (
    takeown /f "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\Office14"
    attrib -r -s -h /s /d "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\Office14"
    rmdir /s /q "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\Office14"
)

if exist "%COMMONPROGRAMFILES%\Microsoft Shared\Office15" (
    takeown /f "%COMMONPROGRAMFILES%\Microsoft Shared\Office15"
    attrib -r -s -h /s /d "%COMMONPROGRAMFILES%\Microsoft Shared\Office15"
    rmdir /s /q "%COMMONPROGRAMFILES%\Microsoft Shared\Office15"
) else if exist "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\Office15" (
    takeown /f "%COMMONPROGRAMFILES%\Microsoft Shared\Office15"
    attrib -r -s -h /s /d "%COMMONPROGRAMFILES%\Microsoft Shared\Office15"
    rmdir /s /q "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\Office15"
)

if exist "%COMMONPROGRAMFILES%\Microsoft Shared\Source Engine" (
    takeown /f "%COMMONPROGRAMFILES%\Microsoft Shared\Source Engine"
    attrib -r -s -h /s /d "%COMMONPROGRAMFILES%\Microsoft Shared\Source Engine"
    rmdir /s /q "%COMMONPROGRAMFILES%\Microsoft Shared\Source Engine"
) else if exist "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\Source Engine" (
    takeown /f "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\Source Engine"
    attrib -r -s -h /s /d "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\Source Engine"
    rmdir /s /q "%COMMONPROGRAMFILES(x86)%\Microsoft Shared\Source Engine"
)

if exist "%APPDATA%\microsoft\templates\Normal.dotm" (
    takeown /f "%APPDATA%\microsoft\templates\Normal.dotm"
    attrib -r -s -h /s /d "%APPDATA%\microsoft\templates\Normal.dotm"
    del /s /q /f "%APPDATA%\microsoft\templates\Normal.dotm"
)

if exist "%APPDATA%\microsoft\templates\Word.dotx" (
    takeown /f "%APPDATA%\microsoft\templates\Word.dotx"
    attrib -r -s -h /s /d "%APPDATA%\microsoft\templates\Word.dotx"
    del /s /q /f "%APPDATA%\microsoft\templates\Word.dotx"
)

if exist "%APPDATA%\microsoft\document building blocks\3082\15\Building Blocks.dotx" (
    takeown /f "%APPDATA%\microsoft\document building blocks\3082\15\Building Blocks.dotx"
    attrib -r -s -h /s /d "%APPDATA%\microsoft\document building blocks\3082\15\Building Blocks.dotx"
    del /s /q /f "%APPDATA%\microsoft\document building blocks\3082\15\Building Blocks.dotx"
)

if exist "%ALLUSERSPROFILE%\Microsoft\Office" (
    takeown /f "%ALLUSERSPROFILE%\Microsoft\Office"
    attrib -r -s -h /s /d "%ALLUSERSPROFILE%\Microsoft\Office"
    rmdir /s /q "%ALLUSERSPROFILE%\Microsoft\Office"
)

if exist "%USERPROFILE%\Microsoft\Office" (
    takeown /f "%USERPROFILE%\Microsoft\Office"
    attrib -r -s -h /s /d "%USERPROFILE%\Microsoft\Office"
    rmdir /s /q "%USERPROFILE%\Microsoft\Office"
)

reg delete "HKCU\Software\Microsoft\Office\15.0\Registration" /f
reg delete "HKCU\Software\Microsoft\Office\16.0\Registration" /f
reg delete "HKCU\Software\Microsoft\Office\Registration" /f

reg delete "HKLM\SOFTWARE\Microsoft\AppVISV" /f

reg delete "HKLM\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot\Virtual" /f
reg delete "HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot\Virtual" /f
reg delete "HKLM\SOFTWARE\Microsoft\Office\Common\InstallRoot\Virtual" /f

reg delete "HKLM\SOFTWARE\Classes\CLSID\{2027FC3B-CF9D-4ec7-A823-38BA308625CC}" /f

reg delete "HKCU\SOFTWARE\Microsoft\Office\15.0\ClickToRun" /f
reg delete "HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun" /f
reg delete "HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRunStore" /f
reg delete "HKLM\SOFTWARE\Microsoft\Office\16.0\ClickToRun" /f
reg delete "HKLM\SOFTWARE\Microsoft\Office\16.0\ClickToRun" /f
reg delete "HKLM\SOFTWARE\Microsoft\Office\16.0\ClickToRunStore" /f
reg delete "HKCU\SOFTWARE\Microsoft\Office\ClickToRun" /f
reg delete "HKLM\SOFTWARE\Microsoft\Office\ClickToRun" /f
reg delete "HKLM\SOFTWARE\Microsoft\Office\ClickToRunStore" /f

reg delete "HKLM\Software\Microsoft\Office\15.0" /f
reg delete "HKLM\Software\Microsoft\Office\15.0" /f
reg delete "HKLM\Software\Microsoft\Office\16.0" /f
reg delete "HKLM\Software\Microsoft\Office\16.0" /f
reg delete "HKLM\SOFTWARE\Microsoft\Office" /f

reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\Lync15" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\Lync16" /f

reg delete "HKLM\SOFTWARE\Classes\Protocols\Handler\osf" /f

reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 1 (ErrorConflict)" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 2 (SyncInProgress)" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 3 (InSync)" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 1 (ErrorConflict)" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 2 (SyncInProgress)" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 3 (InSync)" /f

reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\{B28AA736-876B-46DA-B3A8-84C5E30BA492}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\{8B02D659-EBBB-43D7-9BBA-52CF22C5B025}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\{0875DCB6-C686-4243-9432-ADCCF0B9F2D7}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\{42042206-2D85-11D3-8CFF-005004838597}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\{993BE281-6695-4BA5-8A2A-7AACBFAAB69E}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\{C41662BB-1FA0-4CE0-8DC5-9B7F8279FF97}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\{506F4668-F13E-4AA1-BB04-B43203AB3CC0}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\{D66DC78C-4F61-447F-942B-3FB6980118CF}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\{46137B78-0EC3-426D-8B89-FF7C3A458B5E}" /f

reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{31D09BA0-12F5-4CCE-BE8A-2923E76605DA}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{B4F3A835-0E21-4959-BA22-42B3008E02FF}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{31D09BA0-12F5-4CCE-BE8A-2923E76605DA}" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{B4F3A835-0E21-4959-BA22-42B3008E02FF}" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}" /f

reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{0875DCB6-C686-4243-9432-ADCCF0B9F2D7}" /f

reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\Namespace\{B28AA736-876B-46DA-B3A8-84C5E30BA492}" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\NetworkNeighborhood\Namespace\{46137B78-0EC3-426D-8B89-FF7C3A458B5E}" /f

reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Microsoft Office Temp Files" /f

reg delete "HKCU\Software\Microsoft\Office" /f
reg delete "HKLM\SOFTWARE\Microsoft\Office" /f
reg delete "HKLM\SYSTEM\CurrentControlSet\Services\ose" /f

reg delete "HKLM\SOFTWARE\Microsoft\AppVISV" /f

reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Office" /f

set clave=HKLM\SOFTWARE\Microsoft\Office\Delivery\SourceEngine\Downloads
set valor=0FF1CE}

for /f %%a in ('"reg query "%clave%" | find "%valor%""') do (
    reg delete "%%a" /f
)

set clave=HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\Delivery\SourceEngine\Downloads
set valor=0FF1CE}

for /f %%a in ('"reg query "%clave%" | find "%valor%""') do (
    reg delete "%%a" /f
)

set clave=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
set valor=0FF1CE

for /f %%a in ('"reg query "%clave%" | find "%valor%""') do (
    reg delete "%%a" /f
)

set clave=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
set valor=O365

for /f %%a in ('"reg query "%clave%" | find "%valor%""') do (
    reg delete "%%a - en-us" /f
)

set clave=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
set valor=ProfessionaRetail

for /f %%a in ('"reg query "%clave%" | find "%valor%""') do (
    reg delete "%%a - en-us" /f
)

set clave=HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall
set valor=0FF1CE

for /f %%a in ('"reg query "%clave%" | find "%valor%""') do (
    reg delete "%%a" /f
)

set clave=HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall
set valor=O365

for /f %%a in ('"reg query "%clave%" | find "%valor%""') do (
    reg delete "%%a - en-us" /f
)

set clave=HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall
set valor=ProfessionalRetail

for /f %%a in ('"reg query "%clave%" | find "%valor%""') do (
    reg delete "%%a - en-us" /f
)

reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office14.ENTERPRISE" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office14.ULTIMATE" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office14.PRO" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office14.PROPLUS" /f

reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office14.ENTERPRISER" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office14.ULTIMATER" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office14.PROR" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office14.HOMESTUDENTR" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office14.STANDARDR" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office14.SMALLBUSINESSR" /f

reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ENTERPRISE" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ULTIMATE" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PRO" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PROPLUS" /f

reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ENTERPRISER" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ULTIMATER" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PROR" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\HOMESTUDENTR" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\STANDARDR" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\SMALLBUSINESSR" /f

reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Office14.ENTERPRISE" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Office14.ULTIMATE" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Office14.PRO" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Office14.PROPLUS" /f

reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Office14.ENTERPRISER" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Office14.ULTIMATER" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Office14.PROR" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Office14.HOMESTUDENTR" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Office14.STANDARDR" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Office14.SMALLBUSINESSR" /f

reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ENTERPRISE" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ULTIMATE" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\PRO" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\PROPLUS" /f

reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ENTERPRISER" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ULTIMATER" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\PROR" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\HOMESTUDENTR" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\STANDARDR" /f
reg delete "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\SMALLBUSINESSR" /f

set clave=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products
set valor=F01FEC

for /f %%a in ('"reg query "%clave%" | find "%valor%""') do (
    reg delete "%%a" /f
)

set clave=HKCR\Installer\Features
set valor=F01FEC

for /f %%a in ('"reg query "%clave%" | find "%valor%""') do (
    reg delete "%%a" /f
)

set clave=HKCR\Installer\Products
set valor=F01FEC

for /f %%a in ('"reg query "%clave%" | find "%valor%""') do (
    reg delete "%%a" /f
)

set clave=HKCR\Installer\UpgradeCodes
set valor=F01FEC

for /f %%a in ('"reg query "%clave%" | find "%valor%""') do (
    reg delete "%%a" /f
)

set clave=HKCR\Installer\Win32Assemblies
set valor=Office16

for /f %%a in ('"reg query "%clave%" | find "%valor%""') do (
    reg delete "%%a" /f
)

set clave=HKCR\Installer\Win32Assemblies
set valor=Office15

for /f %%a in ('"reg query "%clave%" | find "%valor%""') do (
    reg delete "%%a" /f
)

set clave=HKCR\Installer\Win32Assemblies
set valor=Office14

for /f %%a in ('"reg query "%clave%" | find "%valor%""') do (
    reg delete "%%a" /f
)

set clave=HKCR\Installer\Win32Assemblies
set valor=Office12

for /f %%a in ('"reg query "%clave%" | find "%valor%""') do (
    reg delete "%%a" /f
)

set clave=HKCR\Installer\Win32Assemblies
set valor=Office11

for /f %%a in ('"reg query "%clave%" | find "%valor%""') do (
    reg delete "%%a" /f
)

reg delete "HKU\.DEFAULT\Software\Microsoft\Office" /f

reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\winword.exe" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\excel.exe" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\powerpnt.exe" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\onenote.exe" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\outlook.exe" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\mspub.exe" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msaccess.exe" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\infopath.exe" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\groove.exe" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\lync.exe" /f

echo Removal Complete!
echo Press any key to exit!
pause >nul 2>&1
endlocal
exit

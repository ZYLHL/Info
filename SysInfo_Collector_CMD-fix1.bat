@echo off
:: SysInfo_Collector_CMD-fix1.bat  v2.2
:: System Information Collector - Universal CMD Edition
:: Compatible: Windows XP SP2 / Vista / 7 / 8 / 8.1 / 10 / 11
::             Windows Server 2003 / 2008 / 2012 / 2016 / 2019 / 2022
::
:: Tier-A (All): systeminfo, wmic, ipconfig, netstat, tasklist, net,
::               reg, tracert, arp, nbtstat, route, pathping, sc, schtasks
:: Tier-B (Vista+): netsh advfirewall, wevtutil, netsh wlan
:: Tier-C (PS 3+):  PowerShell one-liners for enriched data
:: Tier-D (PS 5+):  Get-LocalUser, Get-NetFirewallProfile, winget
::
:: Run as Administrator for full output.
:: Double-click or: SysInfo_Collector_CMD-fix1.bat
:: ============================================================
setlocal EnableDelayedExpansion EnableExtensions

:: ?? 0. Encoding ?????????????????????????????????????????????
chcp 65001 >nul 2>&1

:: ?? 1. Script directory ??????????????????????????????????????
set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

:: ?? 2. Timestamp via WMIC (with robust fallback) ??????????????
set "_DT="
set "TIMESTAMP="
for /f "tokens=2 delims==" %%a in (
    'wmic os get localdatetime /value 2^>nul'
) do set "_DT=%%a"

if defined _DT (
    set "TIMESTAMP=%_DT:~0,4%-%_DT:~4,2%-%_DT:~6,2%_%_DT:~8,2%-%_DT:~10,2%-%_DT:~12,2%"
) else (
    set "TIMESTAMP=%DATE%_%TIME%"
    set "TIMESTAMP=!TIMESTAMP:/=-!"
    set "TIMESTAMP=!TIMESTAMP:\=-!"
    set "TIMESTAMP=!TIMESTAMP::=-!"
    set "TIMESTAMP=!TIMESTAMP: =0!"
)

:: ?? 3. Output file ???????????????????????????????????????????
set "TOTAL=9"
set "OUT=%SCRIPT_DIR%\%COMPUTERNAME%_%TIMESTAMP%.txt"

:: ?? 4. Elevation check (Vista+) ?????????????????????????????
set "IS_ADMIN=0"
net session >nul 2>&1
if %errorlevel%==0 set "IS_ADMIN=1"

:: ?? 5. OS major version via WMIC ????????????????????????????
set "OS_VER=5"
for /f "tokens=2 delims==" %%a in (
    'wmic os get version /value 2^>nul'
) do set "_OSVER=%%a"
for /f "tokens=1 delims=." %%v in ("%_OSVER%") do set "OS_VER=%%v"

:: ?? 6. PowerShell detection ??????????????????????????????????
set "PS_OK=0"
set "PS_VER=0"
powershell.exe -NoProfile -Command "exit 0" >nul 2>&1
if %errorlevel%==0 set "PS_OK=1"
if "%PS_OK%"=="1" (
    for /f "usebackq" %%v in (
        `powershell.exe -NoProfile -Command "$PSVersionTable.PSVersion.Major" 2^>nul`
    ) do set "PS_VER=%%v"
)

:: ?? 7. Console banner ????????????????????????????????????????
cls
echo +----------------------------------------------------------+
echo ^|  SysInfo Collector v2.2  (Universal CMD Edition)        ^|
echo ^|  Compatible: XP / Vista / 7 / 8 / 10 / 11              ^|
echo +----------------------------------------------------------+
echo   Computer : %COMPUTERNAME%
echo   OS Major : %OS_VER%    PS Version: %PS_VER%    Admin: %IS_ADMIN%
echo   Output   : %OUT%
echo.

:: ?? 8. Create output file (no BOM - plain ASCII/UTF-8) ???????
echo ============================================================== > "%OUT%"
echo   SYSTEM INFORMATION REPORT  v2.2  (CMD Universal Edition) >> "%OUT%"
echo ============================================================== >> "%OUT%"
echo Generated   : %DATE% %TIME% >> "%OUT%"
echo Computer    : %COMPUTERNAME% >> "%OUT%"
echo User        : %USERDOMAIN%\%USERNAME% >> "%OUT%"
echo Script      : %~f0 >> "%OUT%"
echo OS Major    : %OS_VER%   PS Available: %PS_OK% (v%PS_VER%) >> "%OUT%"
echo Admin       : %IS_ADMIN% >> "%OUT%"
echo. >> "%OUT%"

:: ============================================================
:: [1/9]  OS and System Basics
:: ============================================================
call :SectionHeader "OS and System Basics" 1
echo [1/%TOTAL%] Collecting OS and system basics...

call :SubHeader "systeminfo"
systeminfo 2>nul >> "%OUT%"

call :SubHeader "OS details (WMIC)"
wmic os get Caption,Version,BuildNumber,OSArchitecture,InstallDate,LastBootUpTime,Manufacturer,RegisteredUser,SerialNumber,WindowsDirectory /format:list 2>nul >> "%OUT%"

call :SubHeader "Computer system (WMIC)"
wmic computersystem get Manufacturer,Model,SystemType,TotalPhysicalMemory,DNSHostName,Domain,Workgroup,UserName /format:list 2>nul >> "%OUT%"

call :SubHeader "BIOS (WMIC)"
wmic bios get Manufacturer,Name,Version,ReleaseDate,SerialNumber,SMBIOSBIOSVersion /format:list 2>nul >> "%OUT%"

call :SubHeader "Timezone (WMIC)"
wmic timezone get Caption,Bias,StandardName /format:list 2>nul >> "%OUT%"

if "%PS_OK%"=="1" (
    call :SubHeader "Uptime (PowerShell)"
    powershell.exe -NoProfile -Command "$u=(Get-Date)-(gcim Win32_OperatingSystem -EA SilentlyContinue).LastBootUpTime; if($u){'Uptime: '+$u.Days+'d '+$u.Hours+'h '+$u.Minutes+'m'}" >> "%OUT%" 2>nul
)

:: ============================================================
:: [2/9]  Hardware
:: ============================================================
call :SectionHeader "Hardware" 2
echo [2/%TOTAL%] Collecting hardware info...

call :SubHeader "CPU (WMIC)"
wmic cpu get Name,NumberOfCores,NumberOfLogicalProcessors,MaxClockSpeed,CurrentClockSpeed,LoadPercentage,SocketDesignation,Manufacturer /format:list 2>nul >> "%OUT%"

call :SubHeader "Physical memory slots (WMIC)"
wmic memorychip get BankLabel,Manufacturer,PartNumber,Capacity,Speed,MemoryType,FormFactor /format:list 2>nul >> "%OUT%"

call :SubHeader "Memory summary (WMIC)"
wmic os get TotalVisibleMemorySize,FreePhysicalMemory,TotalVirtualMemorySize,FreeVirtualMemory /format:list 2>nul >> "%OUT%"

call :SubHeader "Logical disks (WMIC)"
wmic logicaldisk get DeviceID,DriveType,FileSystem,Size,FreeSpace,VolumeName,VolumeSerialNumber /format:list 2>nul >> "%OUT%"

call :SubHeader "Physical disk drives (WMIC)"
wmic diskdrive get Model,InterfaceType,MediaType,Size,Partitions,Status,SerialNumber /format:list 2>nul >> "%OUT%"

call :SubHeader "GPU (WMIC)"
wmic path Win32_VideoController get Name,DriverVersion,VideoModeDescription,AdapterRAM,VideoProcessor,CurrentRefreshRate /format:list 2>nul >> "%OUT%"

call :SubHeader "Monitors (WMIC)"
wmic desktopmonitor get Name,MonitorManufacturer,MonitorType,ScreenHeight,ScreenWidth /format:list 2>nul >> "%OUT%"

call :SubHeader "Sound devices (WMIC)"
wmic sounddev get Name,Manufacturer,Status /format:list 2>nul >> "%OUT%"

if "%PS_VER%" GEQ "3" (
    call :SubHeader "Physical disk health (Get-PhysicalDisk, Win8+)"
    powershell.exe -NoProfile -Command "Get-PhysicalDisk -EA SilentlyContinue | Select FriendlyName,MediaType,BusType,@{N='Size_GB';E={[math]::Round($_.Size/1GB,2)}},HealthStatus,OperationalStatus | Format-List | Out-String -Width 180" >> "%OUT%" 2>nul
)

:: ============================================================
:: [3/9]  Network Configuration
:: ============================================================
call :SectionHeader "Network Configuration" 3
echo [3/%TOTAL%] Collecting network configuration...

call :SubHeader "ipconfig /all"
ipconfig /all 2>nul >> "%OUT%"

call :SubHeader "Network adapters (WMIC)"
wmic nic get Name,MACAddress,NetConnectionID,NetConnectionStatus,Speed,Manufacturer,AdapterType /format:list 2>nul >> "%OUT%"

call :SubHeader "IP-enabled adapter config (WMIC)"
wmic nicconfig where "IPEnabled=TRUE" get Description,IPAddress,IPSubnet,DefaultIPGateway,DNSServerSearchOrder,DHCPEnabled,DHCPServer,MACAddress /format:list 2>nul >> "%OUT%"

call :SubHeader "ARP cache"
arp -a 2>nul >> "%OUT%"

call :SubHeader "NetBIOS local names (nbtstat -n)"
nbtstat -n 2>nul >> "%OUT%"

call :SubHeader "Hosts file (non-comment entries)"
findstr /v "^#" "%SystemRoot%\System32\drivers\etc\hosts" 2>nul | findstr /v "^$" >> "%OUT%"

call :SubHeader "DNS client cache (ipconfig /displaydns)"
ipconfig /displaydns 2>nul >> "%OUT%"

if %OS_VER% GEQ 6 (
    call :SubHeader "Wireless interfaces (netsh wlan, Vista+)"
    netsh wlan show interfaces 2>nul >> "%OUT%"
    call :SubHeader "Wireless profiles (netsh wlan, Vista+)"
    netsh wlan show profiles 2>nul >> "%OUT%"
)

if "%PS_OK%"=="1" (
    call :SubHeader "Public egress IP (PowerShell)"
    powershell.exe -NoProfile -Command "try{$r=(Invoke-RestMethod 'https://api.ipify.org?format=json' -TimeoutSec 5).ip;'Public IP: '+$r}catch{'Public IP: Unavailable'}" >> "%OUT%" 2>nul
)

:: ============================================================
:: [4/9]  Ports and Connections
:: ============================================================
call :SectionHeader "Ports and Connections" 4
echo [4/%TOTAL%] Collecting port and connection info...

call :SubHeader "netstat -ano (all with PID)"
netstat -ano 2>nul >> "%OUT%"

call :SubHeader "netstat -r (routing table)"
netstat -r 2>nul >> "%OUT%"

call :SubHeader "netstat -s (protocol statistics)"
netstat -s 2>nul >> "%OUT%"

if "%PS_VER%" GEQ "3" (
    call :SubHeader "TCP ESTABLISHED with process name (PS Win8+)"
    powershell.exe -NoProfile -Command "Get-NetTCPConnection -State Established -EA SilentlyContinue | Select LocalAddress,LocalPort,RemoteAddress,RemotePort,@{N='Process';E={(Get-Process -Id $_.OwningProcess -EA SilentlyContinue).Name}} | Sort LocalPort | Format-Table -AutoSize | Out-String -Width 200" >> "%OUT%" 2>nul
    call :SubHeader "TCP LISTEN with process name (PS Win8+)"
    powershell.exe -NoProfile -Command "Get-NetTCPConnection -State Listen -EA SilentlyContinue | Select LocalAddress,LocalPort,@{N='Process';E={(Get-Process -Id $_.OwningProcess -EA SilentlyContinue).Name}} | Sort LocalPort | Format-Table -AutoSize | Out-String -Width 200" >> "%OUT%" 2>nul
)

:: ============================================================
:: [5/9]  Process List
:: ============================================================
call :SectionHeader "Process List" 5
echo [5/%TOTAL%] Collecting process list...

call :SubHeader "tasklist /v (verbose)"
tasklist /v 2>nul >> "%OUT%"

call :SubHeader "tasklist /svc (service associations)"
tasklist /svc 2>nul >> "%OUT%"

call :SubHeader "Process details (WMIC)"
wmic process get ProcessId,Name,CommandLine,WorkingSetSize,ThreadCount,HandleCount,ExecutablePath /format:list 2>nul >> "%OUT%"

if "%PS_VER%" GEQ "3" (
    call :SubHeader "Top 40 processes by memory (PS Win8+)"
    powershell.exe -NoProfile -Command "Get-Process | Select Id,Name,@{N='CPU_s';E={[math]::Round($_.CPU,2)}},@{N='Mem_MB';E={[math]::Round($_.WS/1MB,2)}},@{N='Threads';E={$_.Threads.Count}},@{N='Handles';E={$_.HandleCount}},StartTime | Sort Mem_MB -Desc | Select -First 40 | Format-Table -AutoSize | Out-String -Width 250" >> "%OUT%" 2>nul
)

:: ============================================================
:: [6/9]  Startup Items and Services
:: ============================================================
call :SectionHeader "Startup Items and Services" 6
echo [6/%TOTAL%] Collecting startup and services...

call :SubHeader "Startup programs (WMIC)"
wmic startup get Caption,Command,Location,User /format:list 2>nul >> "%OUT%"

call :SubHeader "HKLM Run registry key"
reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" 2>nul >> "%OUT%"

call :SubHeader "HKCU Run registry key"
reg query "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" 2>nul >> "%OUT%"

call :SubHeader "HKLM RunOnce registry key"
reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce" 2>nul >> "%OUT%"

call :SubHeader "All services (sc query)"
sc query type= all state= all 2>nul >> "%OUT%"

call :SubHeader "Service details (WMIC)"
wmic service get Name,DisplayName,State,StartMode,PathName,ProcessId,StartName /format:list 2>nul >> "%OUT%"

call :SubHeader "Scheduled tasks (schtasks)"
schtasks /query /fo LIST /v 2>nul >> "%OUT%"

:: ============================================================
:: [7/9]  Security and Patches
:: ============================================================
call :SectionHeader "Security and Patches" 7
echo [7/%TOTAL%] Collecting security info...

call :SubHeader "Installed hotfixes (wmic qfe)"
wmic qfe get HotFixID,Description,InstalledOn,InstalledBy /format:list 2>nul >> "%OUT%"

call :SubHeader "Local users list (net user)"
net user 2>nul >> "%OUT%"

call :SubHeader "Local groups (net localgroup)"
net localgroup 2>nul >> "%OUT%"

call :SubHeader "Administrators group members"
net localgroup Administrators 2>nul >> "%OUT%"

call :SubHeader "Shared resources (net share)"
net share 2>nul >> "%OUT%"

call :SubHeader "Shared resources detail (WMIC)"
wmic share get Name,Path,Description,Status /format:list 2>nul >> "%OUT%"

call :SubHeader "Open sessions (net session)"
net session 2>nul >> "%OUT%"

call :SubHeader "Firewall status"
if %OS_VER% GEQ 6 (
    netsh advfirewall show allprofiles 2>nul >> "%OUT%"
) else (
    netsh firewall show config 2>nul >> "%OUT%"
)

if %OS_VER% GEQ 6 (
    call :SubHeader "Recent logon events (wevtutil EventID 4624, last 10)"
    wevtutil qe Security /c:10 /rd:true /f:text /q:"*[System[EventID=4624]]" 2>nul >> "%OUT%"
    call :SubHeader "Recent system errors (wevtutil Level=2, last 10)"
    wevtutil qe System /c:10 /rd:true /f:text /q:"*[System[Level=2]]" 2>nul >> "%OUT%"
)

if "%PS_VER%" GEQ "5" (
    call :SubHeader "Local users detail (Get-LocalUser, PS5+)"
    powershell.exe -NoProfile -Command "Get-LocalUser -EA SilentlyContinue | Select Name,Enabled,PasswordRequired,LastLogon,PasswordLastSet,AccountExpires,Description | Format-Table -AutoSize | Out-String -Width 200" >> "%OUT%" 2>nul
    call :SubHeader "Firewall profiles (Get-NetFirewallProfile, PS5+)"
    powershell.exe -NoProfile -Command "Get-NetFirewallProfile -EA SilentlyContinue | Select Name,Enabled,DefaultInboundAction,DefaultOutboundAction | Format-Table -AutoSize" >> "%OUT%" 2>nul
)

:: ============================================================
:: [8/9]  Installed Software
:: ============================================================
call :SectionHeader "Installed Software" 8
echo [8/%TOTAL%] Collecting installed software...

call :SubHeader "Installed products (WMIC - may be slow)"
wmic product get Name,Version,Vendor,InstallDate,InstallLocation /format:list 2>nul >> "%OUT%"

call :SubHeader "Uninstall keys HKLM 64-bit"
reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" /s 2>nul | findstr /i "DisplayName DisplayVersion Publisher InstallDate" >> "%OUT%"

call :SubHeader "Uninstall keys HKLM 32-bit (WOW6432Node)"
reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" /s 2>nul | findstr /i "DisplayName DisplayVersion Publisher InstallDate" >> "%OUT%"

call :SubHeader "Uninstall keys HKCU"
reg query "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" /s 2>nul | findstr /i "DisplayName DisplayVersion Publisher InstallDate" >> "%OUT%"

call :SubHeader "Environment variables (set)"
set 2>nul >> "%OUT%"

call :SubHeader "PATH entries"
for %%p in ("%PATH:;=" "%") do echo   %%~p >> "%OUT%"

if %OS_VER% GEQ 10 (
    call :SubHeader "winget package list (Win10+ best-effort)"
    winget list 2>nul >> "%OUT%"
    if %errorlevel% NEQ 0 echo   winget not available or insufficient privileges >> "%OUT%"
)

:: ============================================================
:: [9/9]  Route Trace
:: ============================================================
call :SectionHeader "Route Trace" 9
set "TRACE_TARGET=61.139.2.69"
echo [9/%TOTAL%] Tracing route to %TRACE_TARGET% -- up to ~90s, please wait...

call :SubHeader "ping -n 4 (connectivity test)"
ping -n 4 %TRACE_TARGET% 2>nul >> "%OUT%"

call :SubHeader "tracert -d (route trace, no DNS lookup)"
tracert -d %TRACE_TARGET% 2>nul >> "%OUT%"

call :SubHeader "route print (full routing table)"
route print 2>nul >> "%OUT%"

call :SubHeader "pathping -n (per-hop packet loss -- may take 3min)"
pathping -n %TRACE_TARGET% 2>nul >> "%OUT%"

:: ============================================================
:: Footer
:: ============================================================
echo. >> "%OUT%"
echo ============================================================== >> "%OUT%"
echo   Collection complete : %DATE% %TIME% >> "%OUT%"
echo   Output file        : %OUT% >> "%OUT%"
echo ============================================================== >> "%OUT%"

echo.
echo +----------------------------------------------------------+
echo ^|   Collection complete!                                  ^|
echo +----------------------------------------------------------+
echo   File saved to:
echo   %OUT%
echo.
explorer /select,"%OUT%"
pause
goto :eof

:: ============================================================
:: Subroutines
:: ============================================================

:SectionHeader
echo. >> "%OUT%"
echo ============================================================== >> "%OUT%"
echo   [%~2/%TOTAL%]  %~1 >> "%OUT%"
echo ============================================================== >> "%OUT%"
echo. >> "%OUT%"
echo. 
echo [%~2/%TOTAL%] %~1
goto :eof

:SubHeader
echo. >> "%OUT%"
echo -- %~1 -- >> "%OUT%"
goto :eof

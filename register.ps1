# Register the Axon Outlook add-in for the current user (no admin). Run after build.ps1, then
# restart Outlook. (For end users, the installer does this automatically.)
$ErrorActionPreference = "Stop"
$dll = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "AxonAddin.dll"
if (-not (Test-Path $dll)) { throw "AxonAddin.dll not found - run build.ps1 first." }

$clsid  = "{7B2C9E14-6A3D-4F58-9C21-3E5A1B7D4F60}"
$progid = "Axon.OutlookAddin"
$asm    = "AxonAddin, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null"
$code   = "file:///" + ($dll -replace '\\', '/')

New-Item -Path "HKCU:\Software\Classes\$progid\CLSID" -Force | Out-Null
Set-ItemProperty "HKCU:\Software\Classes\$progid\CLSID" "(default)" $clsid

$inproc = "HKCU:\Software\Classes\CLSID\$clsid\InprocServer32"
New-Item -Path $inproc -Force | Out-Null
Set-ItemProperty $inproc "(default)" "$env:WINDIR\System32\mscoree.dll"
Set-ItemProperty $inproc "ThreadingModel" "Both"
Set-ItemProperty $inproc "Class" "Axon.OutlookAddin.Connect"
Set-ItemProperty $inproc "Assembly" $asm
Set-ItemProperty $inproc "RuntimeVersion" "v4.0.30319"
Set-ItemProperty $inproc "CodeBase" $code
New-Item -Path "$inproc\0.0.0.0" -Force | Out-Null
Set-ItemProperty "$inproc\0.0.0.0" "Class" "Axon.OutlookAddin.Connect"
Set-ItemProperty "$inproc\0.0.0.0" "Assembly" $asm
Set-ItemProperty "$inproc\0.0.0.0" "RuntimeVersion" "v4.0.30319"
Set-ItemProperty "$inproc\0.0.0.0" "CodeBase" $code

New-Item -Path "HKCU:\Software\Classes\CLSID\$clsid\ProgId" -Force | Out-Null
Set-ItemProperty "HKCU:\Software\Classes\CLSID\$clsid\ProgId" "(default)" $progid

$addins = "HKCU:\Software\Microsoft\Office\Outlook\AddIns\$progid"
New-Item -Path $addins -Force | Out-Null
Set-ItemProperty $addins "FriendlyName" "Axon intelligence"
Set-ItemProperty $addins "Description" "File and download emails with Axon intelligence"
Set-ItemProperty $addins "LoadBehavior" 3 -Type DWord

Write-Host "Registered Axon Outlook add-in. Restart Outlook to load it." -ForegroundColor Green

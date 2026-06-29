# Build the Axon Outlook add-in DLL (managed COM add-in, .NET Framework 4.8).
#   powershell -ExecutionPolicy Bypass -File build.ps1
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

# Outlook locks the loaded DLL — close it so we can recompile.
try { ([Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')).Quit() } catch {}
Start-Sleep 3
Get-Process OUTLOOK -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 1

$csc = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$ext = (Get-ChildItem "C:\Windows\assembly\GAC\Extensibility" -Recurse -Filter "extensibility.dll" -EA SilentlyContinue | Select-Object -First 1).FullName
$off = (Get-ChildItem "C:\Windows\assembly\GAC_MSIL\office"    -Recurse -Filter "OFFICE.DLL"       -EA SilentlyContinue | Select-Object -First 1).FullName
$std = (Get-ChildItem "C:\Windows\assembly\GAC\stdole"         -Recurse -Filter "stdole.dll"       -EA SilentlyContinue | Select-Object -First 1).FullName
if (-not ($ext -and $off -and $std)) { throw "Office interop assemblies not found in the GAC (need Outlook installed)." }

& $csc /nologo /target:library /out:"$root\AxonAddin.dll" /link:"$ext" /link:"$off" /link:"$std" `
    /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:System.Web.Extensions.dll `
    /reference:Microsoft.CSharp.dll /reference:System.dll /reference:System.Net.Http.dll `
    "$root\src\AxonAddin.cs"
if ($LASTEXITCODE -ne 0) { throw "Add-in compile failed." }
Write-Host "Built $root\AxonAddin.dll" -ForegroundColor Green

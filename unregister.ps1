# Remove the Axon Outlook add-in registration for the current user.
$clsid  = "{7B2C9E14-6A3D-4F58-9C21-3E5A1B7D4F60}"
$progid = "Axon.OutlookAddin"
Remove-Item "HKCU:\Software\Classes\CLSID\$clsid" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "HKCU:\Software\Classes\$progid" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "HKCU:\Software\Microsoft\Office\Outlook\AddIns\$progid" -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "Unregistered. Restart Outlook." -ForegroundColor Green

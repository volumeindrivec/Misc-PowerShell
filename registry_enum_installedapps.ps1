$InitialStartMode = Get-CimInstance -ClassName Win32_Service -Filter "Name = 'RemoteRegistry'" | Select-Object -ExpandProperty StartMode

# Get-CimInstance -ClassName Win32_Service -Filter "Name = 'RemoteRegistry'"
Get-Service RemoteRegistry | Set-Service -StartupType Manual | Start-Service

$CurrentStartMode = Get-CimInstance -ClassName Win32_Service -Filter "Name = 'RemoteRegistry'" | Select-Object -ExpandProperty StartMode

Write-Output "Original Start Mode :   $InitialStartMode"
Write-Output "Current Start Mode  :   $CurrentStartMode"

$srv = 'localhost'
$uninstallkey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
$type = [Microsoft.Win32.RegistryHive]::LocalMachine
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($type, $Srv)
$regkey = $reg.OpenSubKey($uninstallkey)

$subkeys = $regKey.GetSubKeyNames()

foreach ($key in $subkeys){
    $thisKey = $uninstallkey+"\\"+$key
    $thisSubKey = $reg.OpenSubKey($thiskey)
    $displayName = $thisSubKey.GetValue("DisplayName")
    Write-Host $displayName

}

#Get-CimInstance -ClassName Win32_Service -Filter "Name = 'RemoteRegistry'"
Get-Service RemoteRegistry | Stop-Service 
Get-Service RemoteRegistry | Set-Service -StartupType $InitialStartMode
#Get-CimInstance -ClassName Win32_Service -Filter "Name = 'RemoteRegistry'"

$CurrentStartMode = Get-CimInstance -ClassName Win32_Service -Filter "Name = 'RemoteRegistry'" | Select-Object -ExpandProperty StartMode

Write-Output "Original Start Mode :   $InitialStartMode"
Write-Output "Current Start Mode  :   $CurrentStartMode"
# Helper function
function Write-Log{
    [cmdletbinding()]
    param(
        [parameter(Mandatory=$true)]$Message,
        $LogPath = "$env:HOMEDRIVE\Logs",
        $LogFile = "$(Split-Path $MyInvocation.ScriptName -Leaf)-$(Get-Date -Format yyyy-MM-dd).log"
    )
    
    if ( -not (Test-Path -Path $LogPath) ){ 
        Write-Output "Path does not exist.  Creating $LogPath"
        New-Item -Path $LogPath -ItemType Directory
    } # End If statement

    Out-File -FilePath $LogPath\$LogFile -Append -NoClobber -InputObject ((Get-Date -Format s) + " - $Message" )
}
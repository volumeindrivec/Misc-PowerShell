## TODO ##
# - Input checking
# - Document (derp)


<#
.SYNOPSIS
    Gets drive space on specified computers using PSRemoting.
.DESCRIPTION
    Gets drive space information using PSRemoting or straight
    WMI calls, depending on specified switches.
.PARAMETER V2
    Uses WMI instead of PSRemoting to be PowerShell V2 compatible.
    Under the hood, it's using Get-WmiObject versus Get-CimInstance.
.EXAMPLE
    Get-DriveSpaceReport -ComputerName localhost
    Simple drive space check using only computer name.
.EXAMPLE
    Get-DriveSpaceReport -ComputerName localhost -DriveType 2
    Check drive space using computer name and specifying a DriveType of 2.
.EXAMPLE
    Get-DriveSpaceReport -ComputerName dc2 -V2
    Check using PowerShell V2 compatibility.
.NOTES
    Version                 :  0.3
    Author                  :  @sukotto_san
    Disclaimer              :  If you run it, you take all responsibility for it.
#>
function Get-DriveSpaceReport {

[CmdletBinding()]
param(
    [parameter(mandatory=$true)][string[]]$ComputerName,
    [int]$DriveType=3,
    [switch]$V2=$false
)

Begin {

    Write-Verbose "Begin Block"

    foreach ($c in $ComputerName) {
        Write-Verbose "Computer  :  $c"
    }
    
}

Process {
    # Enumerate each computer in $ComputerName and get the required info.
    Write-Verbose "Process Block"
    foreach ($computer in $ComputerName) {
        Write-Verbose "Processing $computer"
        try{
            if ( $V2 -eq $True ) {
                Write-Verbose "Using Get-WmiObject calls; -V2 switch was used."
                $os = Get-WmiObject -ComputerName $computer -Class Win32_OperatingSystem -ErrorAction Stop -ErrorVariable OSError
                $disk = Get-WmiObject -ComputerName $computer -Class Win32_LogicalDisk -Filter "drivetype=$DriveType" -ErrorAction Stop -ErrorVariable DiskError
            }
            else {
                Write-Verbose "Using Get-CimInstance via PSRemoting; -V2 switch was NOT used."
                $os = Invoke-Command -ComputerName $computer -ScriptBlock { Get-CimInstance -ClassName Win32_OperatingSystem } -ErrorAction Stop -ErrorVariable OSError
                $disk = Invoke-Command -ComputerName $computer -ScriptBlock { param($dt) Get-CimInstance -ClassName Win32_LogicalDisk -Filter "drivetype=$dt" } -ArgumentList $DriveType -ErrorAction Stop -ErrorVariable DiskError
            }
       
        # Enumerate each drive in $disk. Specifically, this allows for the details on each drive if a computer has more than one.
        foreach ($drive in $disk){
            Write-Verbose "Processing $drive"
            $prop = [ordered]@{
                'ComputerName' = $computer
                'Drive' = $drive.DeviceID
                'PctFree' = $drive.FreeSpace / $drive.Size
                'FreeGB' = $drive.FreeSpace/1GB
                'SizeGB' = $drive.Size/1GB
                'OSName' = $os.Caption
            }
            $object = New-Object -TypeName PSObject -Property $prop
            Write-Output $object
        }
        }
     catch{
            Write-Warning "You done screwed up.  $computer is no con permiso."
     }
    }

}

End { Write-Verbose "End Block" }

}


<#
.SYNOPSIS
    Creates and sends the drive space report to the specified user in HTML formatting.
.DESCRIPTION
    Takes objects generated from the Get-DriveSpaceReport function, formats
    it to HTML, and then sends it to the specified recipient.
.NOTES
    Version                 :  0.1
    Author                  :  @sukotto_san
    Disclaimer              :  If you run it, you take all responsibility for it.
#>
function Send-DriveSpaceReport {

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Mandatory=$True,ValueFromPipeline=$true)]$InputObject,
    [Parameter(Mandatory=$True)][string]$Recipient,
    [Parameter(Mandatory=$True)][string]$Sender,
    [Parameter(Mandatory=$True)][string]$EmailServer,
    $WarningThreshold=30,
    $CriticalThreshold=15
)



Begin{

    Write-Verbose "Begin Block"
    $NormalObjects = @()
    $WarningObjects = @()
    $CriticalObjects = @()

}

Process{
    Write-Verbose "Processing Object: $InputObject"

    if ( $InputObject.PctFree -lt ($CriticalThreshold / 100) ) {
        Write-Verbose "Adding object to Critical Objects"
        $FormattedObject = Select-Object -InputObject $InputObject -Property ComputerName,Drive,@{n="PctFree";e={"{0:P0}" -f $_.PctFree}},@{n="FreeGB";e={"{0:N2}" -f $_.FreeGB}},@{n="SizeGB";e={"{0:N2}" -f $_.SizeGB}}
        $CriticalObjects += $FormattedObject
    }
    elseif ( $InputObject.PctFree -lt ($WarningThreshold / 100) ) {
        Write-Verbose "Adding object to Warning Objects"
        $FormattedObject = Select-Object -InputObject $InputObject -Property ComputerName,Drive,@{n="PctFree";e={"{0:P0}" -f $_.PctFree}},@{n="FreeGB";e={"{0:N2}" -f $_.FreeGB}},@{n="SizeGB";e={"{0:N2}" -f $_.SizeGB}}
        $WarningObjects += $FormattedObject
    }
    else {
        Write-Verbose "Adding object to Normal Object"
        $FormattedObject = Select-Object -InputObject $InputObject -Property ComputerName,Drive,@{n="PctFree";e={"{0:P0}" -f $_.PctFree}},@{n="FreeGB";e={"{0:N2}" -f $_.FreeGB}},@{n="SizeGB";e={"{0:N2}" -f $_.SizeGB}}
        $NormalObjects += $FormattedObject
    }

}

End{
    
    $css = '<style>
            table { width:98%; }
            td { text-align:center; padding:5px; }
            th { background-color:blue; color:white; }
            h3 { text-align:center }
            </style>'

    Write-Verbose "End Block"
    Write-Verbose "Building HTML report"

    $CriticalHTML = $CriticalObjects | ConvertTo-Html -Fragment -PreContent "<h3>CRITICAL - Less than $CriticalThreshold% free</h3>" | Out-String
    $WarningHTML = $WarningObjects | ConvertTo-Html -Fragment -PreContent "<h3>WARNING - Less than $WarningThreshold% free</h3>" | Out-String
    $NormalHTML = $NormalObjects | ConvertTo-Html -Fragment -PreContent "<h3>NORMAL - More than $WarningThreshold% free</h3>" | Out-String

    $Report = ConvertTo-Html -Body "$CriticalHTML $WarningHTML $NormalHTML $css" | Out-String

    Write-Verbose "$Report"
    
    Write-Verbose "Sending Email:
          Recipient   : $Recipient
          Sender      : $Sender
          EmailServer : $EmailServer"

    Send-MailMessage -to $Recipient -From $Sender -Subject "Drive Space Report" -BodyAsHtml $Report -SmtpServer $EmailServer
}

}
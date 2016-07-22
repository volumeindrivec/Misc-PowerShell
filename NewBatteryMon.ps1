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
    Write-Verbose "$(Get-Date -Format s) - $Message"
}

# Helper function
function Get-BatteryChargingStatus { # Not meant to be exported.  Internal to module.
    [CmdletBinding()]
    param()

    Write-Debug "$(Get-Date) - Entering the Get-BatteryChargingStatus function."

    Add-Type -AssemblyName System.Windows.Forms
    $BatteryChargeStatus = [System.Windows.Forms.SystemInformation]::PowerStatus.BatteryChargeStatus.ToString().Split(',').Trim()
    $Charging = $false
  
    foreach ($var in $BatteryChargeStatus){
        Write-Debug "$(Get-Date) - In foreach loop in Get-BatteryChargingStatus function."
        if ($var -eq 'Charging') { $Charging = $true }
    } # End foreach

    return $Charging

} # End Get-BatteryChargingStatus function

function Get-BatteryStatus{
    [CmdletBinding()]
    param()
  
    Add-Type -AssemblyName System.Windows.Forms
    $PowerStatus = [System.Windows.Forms.SystemInformation]::PowerStatus
    $Battery = Get-WmiObject -Namespace root\CIMv2 -Class Win32_Battery
  
    if ($PowerStatus.PowerLineStatus -eq 'Offline'){
        Write-Verbose "$(Get-Date -Format s) - Executing loop to correct for initial bad estimated time when switching to battery."
        Do {
          Start-Sleep -Seconds 1
          $Battery = Get-WmiObject -Namespace root\CIMv2 -Class Win32_Battery
          $EstimatedRuntime = $Battery.EstimatedRuntime
          $EstimatedRuntimeHours = $EstimatedRuntime / 60
          $EstimatedRuntimeHours = '{0:N2}' -f $EstimatedRuntimeHours
          Write-Verbose "$(Get-Date -Format s) - In Get-BatteryStatus Do loop.  `$EstimatedRuntime is $EstimatedRuntime, which is greater than 4320 seconds (3 days)."
        } # End Do block
        While ( ($EstimatedRuntime -gt 4320) -and ($PowerStatus.PowerLineStatus -eq 'Offline') )  #  4320 seconds = 3 days
    } # End If statment

    Write-Verbose "$(Get-Date -Format s) - Creating and writing object."
  
    $props = @{'ComputerName' = $Battery.PSComputerName;
        'Charging' = Get-BatteryChargingStatus;
        'PowerOnline' = $PowerStatus.PowerLineStatus;
        'EstimatedPctRemaining' = $Battery.EstimatedChargeRemaining;
        'EstimatedRuntimeHours' = $EstimatedRuntimeHours;
        'LastAlert' = $null;
        'PowerRestored' = $null
    }

    $obj = New-Object -TypeName PSObject -Property $props
    Write-Output $obj

} # End Get-BatteryStatus function

function Start-BatteryMonitoring {
    [cmdletbinding()]
    
    param(
        [Parameter(Mandatory=$True)][pscredential]$Credential,
        [Parameter(Mandatory=$True)][string[]]$To,
        [Parameter(Mandatory=$True)][string]$From,
        [string]$Subject = "Power alert notification from $env:COMPUTERNAME.",
        [string]$SmtpServer,
        [int]$Port = 25,
        [switch]$UseSsl,
        $AlertThreshold = 30
    )

    Write-Log -Message 'Start-BatteryMonitoring Logging started.'
    # Initialize the hourly writing of a status to the log, so as not to fill the log.  Must be before Do loop.
    $HourlyLog = (Get-Date).AddHours(-1)

    Do {
        # Checks every five seconds.  This is the initial pause.
        Write-Verbose "$(Get-Date -Format s) - Checking status."
        Start-Sleep -Seconds 5
        
        $BatteryStatus = Get-BatteryStatus

        # Battery is not charging and power is not online
        if ((-not $BatteryStatus.Charging) -and ($BatteryStatus.PowerOnline -ne 'Online')){
            $PowerStatusMessage = "Power is $($BatteryStatus.PowerOnline)."
            $ChargeStatusMessage = "Battery charging is $($BatteryStatus.Charging)."
            $PercentRemainingMessage = "Estimated percent remaining is $($BatteryStatus.EstimatedPctRemaining)%."
            $RuntimeStatusMessage = "Estimated runtime is $($BatteryStatus.EstimatedRuntimeHours) hours."
          
            if ( ($BatteryStatus.LastAlert -like $null) -or ((Get-Date) -gt $BatteryStatus.LastAlert.AddMinutes($AlertThreshold)) ) {
            
                Write-Log -Message 'Sending email alert.'
                Write-Log -Message "$PowerStatusMessage"
                Write-Log -Message "$ChargeStatusMessage"
                Write-Log -Message "$PercentRemainingMessage"
                Write-Log -Message "$RuntimeStatusMessage"
    
                $Body = $PowerStatusMessage + "`n" + $ChargeStatusMessage + "`n" + $PercentRemainingMessage + "`n" + $RuntimeStatusMessage + "`n"
          
                Send-MailMessage -To $To -From $From -Subject $Subject -Body $Body -SmtpServer $SmtpServer -Port $Port -UseSsl:$UseSsl -Credential $Credential
                
                $BatteryStatus.PowerRestored = $false
                $BatteryStatus.LastAlert = Get-Date
                $NextAlert = (Get-Date).AddMinutes($AlertThreshold)
         
                Write-Log -Message "Alert sent.  Next alert to be sent at $NextAlert if issue is not corrected."
          
            } # End if block
        
        } # End if block
        # Battery is not charging and power is online (battery maybe not charging due to already full)
        elseif ((-not $BatteryStatus.Charging) -and ($BatteryStatus.PowerOnline -eq 'Online')){ 
            $BatteryStatus.LastAlert = (Get-Date).AddHours(-2)
            $PowerStatusMessage = "Power is $($BatteryStatus.PowerOnline)."
            $ChargeStatusMessage = "Battery charging is $($BatteryStatus.Charging)."
            $PercentRemainingMessage = "Estimated percent remaining is $($BatteryStatus.EstimatedPctRemaining)%."
            $RuntimeStatusMessage = "Estimated runtime is $($BatteryStatus.EstimatedRuntimeHours) hours."

            if (-not $BatteryStatus.PowerRestored) {
                Write-Verbose 'Power online, battery not charging.'
                Write-Log -Message 'Sending power restored alert.'
                Write-Log -Message "$PowerStatusMessage"
                Write-Log -Message "$ChargeStatusMessage"
                Write-Log -Message "$PercentRemainingMessage"
                Write-Log -Message "$RuntimeStatusMessage"

                $Body = $PowerStatusMessage + "`n" + $ChargeStatusMessage + "`n" + $PercentRemainingMessage + "`n" + $RuntimeStatusMessage + "`n"
                Send-MailMessage -To $To -From $From -Subject $Subject -Body $Body -SmtpServer $SmtpServer -Port $Port -UseSsl:$UseSsl -Credential $Credential
          
                Write-Verbose -Message "$BatteryStatus"
                Write-Log -Message 'Power restored alert sent.'
                $BatteryStatus.PowerRestored = $True
                write-verbose -Message "$BatteryStatus"
                start-sleep -Seconds 10
            } # End if block

        } # End ElseIf block
        # Battery charging and power online
        elseif (($BatteryStatus.Charging) -and ($BatteryStatus.PowerOnline -eq 'Online')){ 
            Write-Verbose "$(Get-Date) - Power restored, charging."
            $BatteryStatus.LastAlert = (Get-Date).AddHours(-2)
            $PowerStatusMessage = "Power is $($BatteryStatus.PowerOnline)."
            $ChargeStatusMessage = "Battery charging is $($BatteryStatus.Charging)."
            $PercentRemainingMessage = "Estimated percent remaining is $($BatteryStatus.EstimatedPctRemaining)%."
       
            if (-not $BatteryStatus.PowerRestored) {
                Write-Verbose 'Power Online, Battery Charging'
                Write-Log -Message 'Sending power restored email alert.'
                Write-Log -Message "$PowerStatusMessage"
                Write-Log -Message "$ChargeStatusMessage"
                Write-Log -Message "$PercentRemainingMessage"
                Write-Log -Message "$RuntimeStatusMessage"

                $Body = $PowerStatusMessage + "`n" + $ChargeStatusMessage + "`n" + $PercentRemainingMessage + "`n" + $RuntimeStatusMessage + "`n"
                Send-MailMessage -To $To -From $From -Subject $Subject -Body $Body -SmtpServer $SmtpServer -Port $Port -UseSsl:$UseSsl -Credential $Credential
          
                Write-Log -Message 'Power restored alert sent.'
                $BatteryStatus.PowerRestored = $True
            } # End if block
       
        } # End ElseIf block

        # Write hourly status to the log file.
        if ((Get-Date) -ge $HourlyLog ) {
            Write-Log -Message 'Hourly Status'
            Write-Log -Message "Power is $($BatteryStatus.Charging)."
            Write-Log -Message "Battery charging is $($BatteryStatus.Charging)."
            Write-Log -Message "Estimated percent remaining is $($BatteryStatus.EstimatedPctRemaining)%."
            if ( $BatteryStatus.EstimatedRuntimeHours ) {
                Write-Log -Message "Estimated runtime is $($BatteryStatus.EstimatedRuntimeHours) hours."
            } # End If block
            $HourlyLog = (Get-Date).AddHours(1)
        }
        
    } # End Do block
    While ( -not $StartShutdown )

} # End Start-BatteryMonitoring function


Start-BatteryMonitoring -To scottw@rbogeek.com -From scottw@rbogeek.com -SmtpServer smtp.office365.com -Port 587 -UseSsl -Verbose

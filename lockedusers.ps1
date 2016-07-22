function Get-RecentLockedUsers {
$computerName = $env:COMPUTERNAME
[string[]]$array = $null

Get-EventLog -LogName Security -After ((Get-Date).AddMinutes(-11)) | 
Where-Object -FilterScript { ( $_.EventID -match '4625' ) -and ( $_.Category -match '12546') } |


ForEach-Object { 
    $user = $_.ReplacementStrings[5]
    $array = $array + $user
    }

$array = $array | Select-Object -Unique


if ($array -ne $null){
    $message = "The following users are locked out within the past few minutes: " + $array | Out-String
    Send-MailMessage -From "$computerName@<SERVER>.com" -To '<USER>@<SERVER>.com' -Subject "User(s) locked out report on $computerName" -Body $message -SmtpServer 'SMTPSERVER'
    }

}

Get-RecentLockedUsers

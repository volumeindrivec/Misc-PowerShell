Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin
& 'C:\Program Files\Microsoft\Exchange Server\Bin\Exchange.ps1'

function Get-SMTPAddress {
    
    param (
        $smtpAddress
    )

    Get-Recipient -ResultSize Unlimited | 
     Select-Object -Property Name -ExpandProperty EmailAddresses | 
     Where-Object -FilterScript { $_.SmtpAddress -like "*$smtpAddress*" } | 
     Select-Object -Property Name,SmtpAddress
}
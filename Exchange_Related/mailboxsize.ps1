Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin -ErrorAction 0

Get-MailboxStatistics | Select-Object -Property DisplayName, DatabaseName, @{ l='MailboxSize (MB)'; e= { $_.TotalItemSize.Value.ToMB() } } | Sort-Object -Property 'MailboxSize (MB)' -Descending
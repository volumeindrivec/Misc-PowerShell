<#

From Chris Brenton's threat hunting class by ACM

*nix command:
  tshark -r thunt-lab.pcapng -T fields -e dns.qry.name | sort | uniq | rev | cut -d '.' -f 1-2 | reve | sort | uniq -c | sort -rn | head -10

PowerShell variant below uses a few variables for reuse, and also eliminates blank lines.
#>

# $Domains = 'C:\Scripts\Domains.txt'
$Domains =  & 'C:\Program Files\Wireshark\tshark.exe' -r 'C:\Scripts\pcaps\out.pcap' -T fields -e dns.qry.name
$SubDomains = 1
$Top = 20

$Domains.Where({ $_ -ne "" }) | Sort-Object | Select-Object -Unique | rev | ForEach-Object { $_.split(".")[0..$SubDomains] -join "." } | rev `
| Sort-Object | Group-Object | Select-Object Count,Name | Sort-Object -Property Count -Descending | Select-Object -First $Top | ft
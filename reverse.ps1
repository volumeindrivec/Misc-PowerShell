# $Domains = 'C:\Scripts\Domains.txt'
$Domains =  & 'C:\Program Files\Wireshark\tshark.exe' -r 'C:\Scripts\dns.pcap' -T fields -e dns.qry.name
$SubDomains = 1
$Top = 20

$Domains.Where({ $_ -ne "" }) | Sort-Object | Select-Object -Unique | rev | ForEach-Object { $_.split(".")[0..$SubDomains] -join "." } | rev `
| Sort-Object | Group-Object | Select-Object Count,Name | Sort-Object -Property Count -Descending | Select-Object -First $Top | ft
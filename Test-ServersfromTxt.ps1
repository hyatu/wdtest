$filePath = "C:\Temp\IPAddress.txt"  # the path with the IP address text file

function Test-ServerOnline {
    param (
        [string]$IPAddress
    )
    
    $ping = New-Object System.Net.NetworkInformation.Ping
    $result = $ping.Send($IPAddress)
    
    return $result.Status -eq 'Success'
}

$ipAddresses = Get-Content -Path $filePath

$results = foreach ($ip in $ipAddresses) {
    if (Test-ServerOnline -IPAddress $ip) {
        try {
            $hostEntry = [System.Net.Dns]::GetHostEntry($ip)
            $dnsName = $hostEntry.HostName
            [PSCustomObject]@{
                IPAddress = $ip
                DNSName = $dnsName
                Status = "Online"
            }
        }
        catch {
            [PSCustomObject]@{
                IPAddress = $ip
                DNSName = "Not Found"
                Status = "Online"
            }
        }
    }
    else {
        [PSCustomObject]@{
            IPAddress = $ip
            DNSName = "N/A"
            Status = "Offline"
        }
    }
}

$results | Export-Csv -Path "C:\Temp\DNSNameStatus.csv" -NoTypeInformation

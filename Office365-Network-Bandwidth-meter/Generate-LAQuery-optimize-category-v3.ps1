<#

.COPYRIGHT
Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
See LICENSE in the project root for license information.

#>

####################################################

#Script to generate Log Analytics query for Office 365 endpoint categories optimize and allow
# Accepted input for Category parameter are Optimize or Allow
# Accepted input format for startdate and enddate is 
#startdate: 2018-11-01T09:00
#enddate: 2018-11-12T09:00

######################################################################################################

[CmdletBinding(SupportsShouldProcess=$True)]

Param (
    
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $Category,
    

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $startdate,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $enddate

      
    )

   


function Get-IPrange
{
<# 
  .SYNOPSIS  
    Get the IP addresses in a range 
  .EXAMPLE 
   Get-IPrange -start 192.168.8.2 -end 192.168.8.20 
  .EXAMPLE 
   Get-IPrange -ip 192.168.8.2 -mask 255.255.255.0 
  .EXAMPLE 
   Get-IPrange -ip 192.168.8.3 -cidr 24 
#> 
 
param 
( 
  [string]$start, 
  [string]$end, 
  [string]$ip, 
  [string]$mask, 
  [int]$cidr 
) 
 
function IP-toINT64 () { 
  param ($ip) 
 
  $octets = $ip.split(".") 
  return [int64]([int64]$octets[0]*16777216 +[int64]$octets[1]*65536 +[int64]$octets[2]*256 +[int64]$octets[3]) 
} 
 
function INT64-toIP() { 
  param ([int64]$int) 

  return (([math]::truncate($int/16777216)).tostring()+"."+([math]::truncate(($int%16777216)/65536)).tostring()+"."+([math]::truncate(($int%65536)/256)).tostring()+"."+([math]::truncate($int%256)).tostring() )
} 
 
if ($ip) {$ipaddr = [Net.IPAddress]::Parse($ip)} 
if ($cidr) {$maskaddr = [Net.IPAddress]::Parse((INT64-toIP -int ([convert]::ToInt64(("1"*$cidr+"0"*(32-$cidr)),2)))) } 
if ($mask) {$maskaddr = [Net.IPAddress]::Parse($mask)} 
if ($ip) {$networkaddr = new-object net.ipaddress ($maskaddr.address -band $ipaddr.address)} 
if ($ip) {$broadcastaddr = new-object net.ipaddress (([system.net.ipaddress]::parse("255.255.255.255").address -bxor $maskaddr.address -bor $networkaddr.address))} 
 
if ($ip) { 
  $startaddr = IP-toINT64 -ip $networkaddr.ipaddresstostring 
  $endaddr = IP-toINT64 -ip $broadcastaddr.ipaddresstostring 
} else { 
  $startaddr = IP-toINT64 -ip $start 
  $endaddr = IP-toINT64 -ip $end 
} 
 
 
INT64-toIP -int $startaddr
INT64-toIP -int $endaddr
}

function GetIpAddressesForCategory ($optimize) {
    <# 
    .SYNOPSIS  
    Get the IP addresses for a given O365 endpoint category
    .EXAMPLE 
    GetIpAddressesForCategory("Optimize")
    #> 

    #if (!$optimize) { 
     #   $optimize = "Optimize"
    #} 
    # webservice root URL
    $ws = "https://endpoints.office.com"
    $clientRequestId = [guid]::NewGuid()
    $O365instance = "Worldwide"


    # invoke endpoints method to get the new data
    $endpointSets = Invoke-RestMethod -Uri ($ws + "/endpoints/"+$O365instance+"?clientRequestId=" + $clientRequestId)
    

    $flatIps = $endpointSets | ForEach-Object {
        $endpointSet = $_
        if ($endpointSet.category -eq $category ) {
            $ips = $(if ($endpointSet.ips.Count -gt 0) { $endpointSet.ips } else { @() })
            # IPv4 strings have dots while IPv6 strings have colons
            $ip4s = $ips | Where-Object { $_ -like '*.*' }
            
            $IpCustomObjects = @()
            if ($endpointSet.tcpPorts -or $endpointSet.udpPorts) {
                $IpCustomObjects = $ip4s | ForEach-Object {
                    [PSCustomObject]@{
                        category = "Allow";
                        ip = $_;
                        tcpPorts = $endpointSet.tcpPorts;
                        udpPorts = $endpointSet.udpPorts;
                    }
                }
            }
            $IpCustomObjects
        } 
        
    }

    return $flatIps
}

function Get-BwData($optimize) {
    $mySubnets = GetIpAddressesForCategory($optimize)
    $firstItem = 0

  $IpRanges = $mySubnets.ip | ForEach-Object {
    
    $mySubnet = $_.split("/")
    
    if ($mySubnet[1] -eq "32") {
        $IpRange = $mySubnet[0]
        if ($firstItem -eq 0) {
            $KustoQuery += " | where (parse_ipv4(DestinationIp) == parse_ipv4('"+$IpRange+"'))"
        } else {
            $KustoQuery += " or (parse_ipv4(DestinationIp) == parse_ipv4('"+$IpRange+"'))"
        }
    } elseif ($mysubnet[1] -ne "32") {
        
        $IpRange = Get-IPrange -ip $mySubnet[0] -cidr $mySubnet[1]
        
        

            if ($firstItem -eq 0) {
                $KustoQuery += " | where (parse_ipv4(DestinationIp) >= parse_ipv4('"+$IpRange[0]+"') and parse_ipv4(DestinationIp) <= parse_ipv4('"+$IpRange[1]+"')) "
            } else {
                $KustoQuery += " or (parse_ipv4(DestinationIp) >= parse_ipv4('"+$IpRange[0]+"') and parse_ipv4(DestinationIp) <= parse_ipv4('"+$IpRange[1]+"')) "
            }

    }
      $firstItem = 1
  } 

  return $KustoQuery
}



$OutputData = "VMConnection " # Change to VMConnection

$secondService = 1
foreach ($optimize in $category) {
  $OutputData += Get-BwData($optimize)
  $secondService++
} 

$OutputData += " | where TimeGenerated > todatetime('"+$startdate+"') and TimeGenerated < todatetime('"+$enddate+"')"

##$OutputData += " | project SessionStartTime, TotalBytes " ## Potentially not needed anymore.

$ReturnMessage = ''+$OutputData+''

$Date = Get-Date 
$Filename = "LogAnalyticsquery" + "_" + "$category" + "_" + $Date.Day + "-" + $Date.Month + "-" + $Date.Year + "_" + $Date.Hour + "-" + $Date.Minute + ".txt"

Out-File -Encoding ascii -FilePath "$ENV:Temp\$FileName" -InputObject $ReturnMessage
Write-Host "Log Analytics query written to '$ENv:Temp\$FileName'" -ForegroundColor Yellow
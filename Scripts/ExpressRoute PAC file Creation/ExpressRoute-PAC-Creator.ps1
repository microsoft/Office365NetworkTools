<#

.COPYRIGHT
Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
See LICENSE in the project root for license information.

#>

####################################################

[CmdletBinding(SupportsShouldProcess=$True)]

Param (
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [String] $ProxySettings,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $TenantName
)

####################################################

# webservice root URL
$baseServiceUrl = "https://endpoints.office.com"

$Date = Get-Date

$clientRequestId = [guid]::NewGuid()

if($TenantName){

    $uri = "$baseServiceUrl" + "/endpoints/worldwide?clientRequestId=$clientRequestId" + "&TenantName=$TenantName"

}

else {

    $uri = "$baseServiceUrl" + "/endpoints/worldwide?clientRequestId=$clientRequestId"

}

# invoke endpoints method to get the new data
$endpointSets = Invoke-RestMethod -Uri ($uri)

Write-Host "Service URL:" -f Yellow
write-host "$uri"

####################################################

$ExpressRoute = $endpointSets | ? { $_.expressRoute -eq "true" }

$ExpressRouteUrls = $ExpressRoute.urls | sort -Unique

if(!($ProxySettings)){ $ProxySettings = 'PROXY 10.10.10.10:8080' }

# PAC File Construction
$PACFile = @"
// JavaScript source code
// Script created $Date
// Every Effort is made to ensure 100% accuracy but this PAC should be used as an example and cross-checked with your traffic flow needs and the Office 365 URL & IP page. 
// Intended only for Worldwide Office 365 instances, which the vast majority of customers will be using
// PAC presumes all Office 365 BGP communities/route filters are allowed.
// As the tenant name is requested upon running this script, we're able to send your tenant specific SharePoint/OneDrive traffic direct rather than *.sharepoint.com which would encompass any Office 365 tenant.
// If you wish to send more than one tenant's SPO/ODfB traffic direct, duplicate the .sharepoint.com URLs and replace the tenant name.

function FindProxyForURL(url, host)
{
    // Define proxy server
    var proxyserver = "$ProxySettings";
    // Make host lowercase
    var lhost = host.toLowerCase();
    host = lhost;
    //SUB-FQDNs of ExpressRoutable wildcards which need to be explicitly sent to the proxy at the top of the PAC because they arent ER routable
    if ((shExpMatch(host, "quicktips.skypeforbusiness.com")))
               				
    {
        return proxyserver;
    }
        //EXPRESS ROUTE DIRECT
    else if ((isPlainHostName(host))

"@

# Looping through all ExpressRoute URLs and adding to PAC File
foreach($URL in $ExpressRouteUrls){

$PACFile += @"
            || (shExpMatch(host, "$URL"))  
            
"@

}

# PAC File End
$PACFile += @"
       )
    {
        return "DIRECT";
    }

        //Catchall for all other traffic to proxy
    else
    {
        return proxyserver;
    }
}
"@

####################################################

$Filename = "ExpressRoutePAC" + "_" + $Date.Day + "-" + $Date.Month + "-" + $Date.Year + "_" + $date.Hour + "-" + $date.Minute + ".txt"

$PACFile | Out-File -Encoding ascii -FilePath "$ENV:Temp\$FileName"

Write-Host
Write-Host "ExpressRoutePAC output written to '$ENv:Temp\$FileName'" -ForegroundColor Gray
Write-Host
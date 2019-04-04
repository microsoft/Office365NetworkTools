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
    [String] $DefaultProxySettings,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [String] $SecondaryProxySettings,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $TenantName,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 2)]
    [int] $Scenario = 1
)

####################################################

# webservice root URL
$baseServiceUrl = "https://endpoints.office.com"

if(!($DefaultProxySettings)){ $DefaultProxySettings = 'PROXY 10.10.10.10:8080' }
if(!($SecondaryProxySettings)){ $SecondaryProxySettings = 'PROXY 20.20.20.20:8080' }

$Date = Get-Date

$clientRequestId = [guid]::NewGuid()

####################################################

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

# PAC File Construction
$PACFile = @"
// JavaScript source code
// Script created $Date
// Every Effort is made to ensure 100% accuracy but this PAC should be used as an example and cross-checked with your traffic flow needs and the Office 365 URL & IP page. 
// Intended only for Worldwide Office 365 instances, which the vast majority of customers will be using
// As the tenant name is requested upon running this script, we're able to send your tenant specific SharePoint/OneDrive traffic direct rather than *.sharepoint.com which would encompass any Office 365 tenant.
// If you wish to send more than one tenant's SPO/ODfB traffic direct, duplicate the .sharepoint.com URLs and replace the tenant name.Or send *.sharepoint.com via the proxy to allow access to any SPO/ODfB tenant.

function FindProxyForURL(url, host)
{
    // Define proxy server
    var proxyserver = "$DefaultProxySettings";
    var proxyserver2 = "$SecondaryProxySettings";
    // Make host lowercase
    var lhost = host.toLowerCase();
    host = lhost;

"@

####################################################

#region Optimize

$O_Urls = ($endpointSets | ? { $_.category -eq "Optimize" }).urls | sort -Unique

# Optimize Direct Contruction
$PACFile += @"

    //OPTIMIZE DIRECT
    if ((isPlainHostName(host))

"@

# Looping through all Optimize URLs and adding to PAC File
foreach($URL in $O_Urls){

$PACFile += @"
            || (shExpMatch(host, "$URL"))

"@

}

$PACFile += @"
       
    )
    {
        return "DIRECT";
    }

"@

#endregion

####################################################

if($Scenario -eq 1){

$AD_Urls = ($endpointSets | ? { $_.category -in @("Allow","Default") }).urls | sort -Unique

$PACFile += @"

    //ALLOW DEFAULT PROXY
    else if ((isPlainHostName(host))

"@

# Looping through all Optimize URLs and adding to PAC File
foreach($URL in $AD_Urls){

$PACFile += @"
            || (shExpMatch(host, "$URL"))  

"@

}

$PACFile += @"
       
    )
    {
        return proxyserver;
    }
"@

}

####################################################

elseif($Scenario -eq 2){

$A_Urls = ($endpointSets | ? { $_.category -eq "Allow" }).urls | sort -Unique
$D_Urls = ($endpointSets | ? { $_.category -eq "Default" }).urls | sort -Unique

$PACFile += @"

    //ALLOW PROXY
    else if ((isPlainHostName(host))

"@

# Looping through all Optimize URLs and adding to PAC File
foreach($URL in $A_Urls){

$PACFile += @"
            || (shExpMatch(host, "$URL"))  

"@

}

$PACFile += @"
       
    )
    {
        return proxyserver;
    }

"@

$PACFile += @"

    //DEFAULT PROXY
    else if ((isPlainHostName(host))

"@

# Looping through all Optimize URLs and adding to PAC File
foreach($URL in $D_Urls){

$PACFile += @"
            || (shExpMatch(host, "$URL"))  

"@

}

$PACFile += @"
       
    )
    {
        return proxyserver2;
    }

"@

}

####################################################

# PAC File End
$PACFile += @"

}
"@

####################################################

$Filename = "o365PAC" + "_" + $Date.Day + "-" + $Date.Month + "-" + $Date.Year + "_" + $date.Hour + "-" + $date.Minute + ".txt"

$PACFile | Out-File -Encoding ascii -FilePath "$ENV:Temp\$FileName"

Write-Host
Write-Host "o365PAC output written to '$ENv:Temp\$FileName'" -ForegroundColor Gray
Write-Host
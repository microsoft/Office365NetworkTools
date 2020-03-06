<#

.COPYRIGHT
Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
See LICENSE in the project root for license information.

#>

####################################################

# webservice root URL
$ws = "https://endpoints.office.com"
$baseServiceUrl = "https://endpoints.office.com"

# path where client ID and latest version number will be stored
$datapath = $Env:TEMP + "\endpoints_clientid_latestversion.txt"

####################################################

# fetch client ID and version if data file exists; otherwise create new file

if (Test-Path $datapath) {

    $content = Get-Content $datapath
    $clientRequestId = $content[0]
    $lastVersion = $content[1]

}

else {

    $clientRequestId = [GUID]::NewGuid().Guid
    $lastVersion = "0000000000"
    #@($clientRequestId, $lastVersion) | Out-File $datapath

}

####################################################

# call version method to check the latest version, and pull new data if version number is different
$version = Invoke-RestMethod -Uri ($ws + "/version?clientRequestId=" + $clientRequestId)

if ($version.latest -gt $lastVersion){

    Write-Host
    Write-Host "New version of Office 365 worldwide commercial service instance endpoints detected" -ForegroundColor Cyan
    Write-Host
    
    # write the new version number to the data file
    #@($clientRequestId, $version.latest) | Out-File $datapath
    
    ####################################################

    Write-Host "If you have Office 365 Sharepoint Online, please specify tenant name (e.g. if your tenant is named contoso.sharepoint.com then write contoso):" -ForegroundColor Yellow
    $TenantName = read-host
    Write-Host

    ####################################################

    #region Optimize

    # invoke endpoints method to get the new data
    
    if($TenantName){

        $uri = "$baseServiceUrl" + "/endpoints/worldwide?clientRequestId=$clientRequestId" + "&TenantName=$TenantName"

    }

    else {

        $uri = "$baseServiceUrl" + "/endpoints/worldwide?clientRequestId=$clientRequestId"

    }

    # invoke endpoints method to get the new data
    $endpointSets = Invoke-RestMethod -Uri ($uri)

    write-host "$uri"
    Write-Host

    $Optimize = $endpointSets | Where-Object { $_.category -eq "Optimize" }

    $OptimizeUrls = $Optimize.urls | Sort-Object -Unique

    $optimizeIpsv4 = $Optimize.ips | Where-Object { ($_).contains(".") } | Sort-Object -Unique
    $optimizeIpsv6 = $Optimize.ips | Where-Object { ($_).contains(":") } | Sort-Object -Unique

    ####################################################

    $optimizeTcpPorts = $Optimize.tcpPorts | Sort-Object -Unique

    ####################################################
    
    $optimizeUdpPorts = $Optimize.udpPorts | Sort-Object -Unique

    #endregion

    ####################################################

    #region Allow

    $Allow = $endpointSets | Where-Object { $_.category -eq "Allow" }

    $AllowUrls = $Allow.urls | Sort-Object -Unique

        $allowUrlsCustomObject = @()

        foreach($allowUrl in $allowUrls) {
            
            if($optimizeUrls -notcontains $allowUrl){
            
                $allowUrlsCustomObject += $allowUrl
            
            }

        }

    ####################################################

    $AllowIps = $Allow.ips | Sort-Object -Unique

    $allowIpsCustomObject = @()

    foreach($allowIP in $allowIps) {
            
        if($optimizeIps -notcontains $allowIP){
            
            $allowIpsCustomObject += $allowIP
            
        }

    }
    
    $AllowIpsv4 = $allowIpsCustomObject | Where-Object { ($_).contains(".") }
    $AllowIpsv6 = $allowIpsCustomObject | Where-Object { ($_).contains(":") }

    ####################################################

    $AllowTcpPorts = $Allow.tcpPorts | Sort-Object -Unique

    ####################################################
    
    $AllowUdpPorts = $Allow.udpPorts | Sort-Object -Unique

    #endregion

    ####################################################

    #region Default

    $Default = $endpointSets | Where-Object { $_.category -eq "Default" }

    $DefaultUrls = $Default.urls | Sort-Object -Unique

        $DefaultUrlsCustomObject = @()

        foreach($DefaultUrl in $DefaultUrls) {
            
            if($AllowUrls -notcontains $DefaultUrl){
            
                $DefaultUrlsCustomObject += $DefaultUrl
            
            }

        }

    ####################################################

    $DefaultIps = $Default.ips | Sort-Object -Unique

    $DefaultIpsCustomObject = @()

    foreach($DefaultIP in $DefaultIps) {
            
        if($AllowIps -notcontains $DefaultIP){
            
            $DefaultIpsCustomObject += $DefaultIP
            
        }

    }
    
    $DefaultIpsv4 = $DefaultIpsCustomObject | Where-Object { ($_).contains(".") }
    $DefaultIpsv6 = $DefaultIpsCustomObject | Where-Object { ($_).contains(":") }

    ####################################################

    $DefaultTcpPorts = $Default.tcpPorts | Sort-Object -Unique

    ####################################################
    
    $DefaultUdpPorts = $Default.udpPorts | Sort-Object -Unique

    #endregion

    ####################################################

    #region Output

    $fileName_Optimize = "o365_Optimize.txt"

    if(Test-Path "$Env:TEMP\$fileName_Optimize"){ 
    
        Set-Content -Value "" "$Env:TEMP\$fileName_Optimize"
        
    }

    $text = "Optimize: URLs to be sent via an optimized path"
    Write-Host $text -ForegroundColor Yellow
    $text | Add-Content "$Env:TEMP\$fileName_Optimize"
    $optimizeUrls
    $optimizeUrls | Add-Content "$Env:TEMP\$fileName_Optimize"
    Write-Host
    "" | Add-Content "$Env:TEMP\$fileName_Optimize"

    $text = "Optimize: IPv4 Address required to be accessible via the optimized path"
    Write-Host $text -ForegroundColor Yellow
    $text | Add-Content "$Env:TEMP\$fileName_Optimize"
    $optimizeIpsv4
    $optimizeIpsv4 | Add-Content "$Env:TEMP\$fileName_Optimize"
    Write-Host
    "" | Add-Content "$Env:TEMP\$fileName_Optimize"

    $text = "Optimize: IPv6 Address required to be accessible via the optimized path (if IPV6 desired)"
    Write-Host $text -ForegroundColor Yellow
    $text | Add-Content "$Env:TEMP\$fileName_Optimize"
    $optimizeIpsv6
    $optimizeIpsv6 | Add-Content "$Env:TEMP\$fileName_Optimize"
    Write-Host
    "" | Add-Content "$Env:TEMP\$fileName_Optimize"

    $text = "Optimize: TCP Ports required to be allowed through the Optimized path"
    Write-Host $text -ForegroundColor Yellow
    $text | Add-Content "$Env:TEMP\$fileName_Optimize"
    $optimizeTcpPorts
    $optimizeTcpPorts | Add-Content "$Env:TEMP\$fileName_Optimize"
    Write-Host
    "" | Add-Content "$Env:TEMP\$fileName_Optimize"

    $text = "Optimize: UDP Ports required to be allowed through the Optimized path"
    Write-Host $text -ForegroundColor Yellow
    $text | Add-Content "$Env:TEMP\$fileName_Optimize"
    $optimizeUdpPorts
    $optimizeUdpPorts | Add-Content "$Env:TEMP\$fileName_Optimize"
    Write-Host
    "" | Add-Content "$Env:TEMP\$fileName_Optimize"

    Write-Host "-----------------------------------------------------"
    Write-Host

    ####################################################

    $fileName_Allow = "o365_Allow.txt"

    if(Test-Path "$Env:TEMP\$fileName_Allow"){ 
    
        Set-Content -Value "" "$Env:TEMP\$fileName_Allow"
        
    }

    $text = "Allow: URLs required for the service to work (Can be proxied if optimized)"
    Write-Host $text -ForegroundColor Yellow
    $text | Add-Content "$Env:TEMP\$fileName_Allow"
    $allowUrlsCustomObject
    $allowUrlsCustomObject | Add-Content "$Env:TEMP\$fileName_Allow"
    Write-Host
    "" | Add-Content "$Env:TEMP\$fileName_Allow"

    $text = "Allow: IPv4 Address required to be allowed"
    Write-Host $text -ForegroundColor Yellow
    $text | Add-Content "$Env:TEMP\$fileName_Allow"
    $AllowIpsv4
    $allowIpsv4 | Add-Content "$Env:TEMP\$fileName_Allow"
    Write-Host
    "" | Add-Content "$Env:TEMP\$fileName_Allow"

    $text = "Allow: IPv6 Address required to be allowed"
    Write-Host $text -ForegroundColor Yellow
    $text | Add-Content "$Env:TEMP\$fileName_Allow"
    $AllowIpsv6
    $allowIpsv6 | Add-Content "$Env:TEMP\$fileName_Allow"
    Write-Host
    "" | Add-Content "$Env:TEMP\$fileName_Allow"

    $text = "Allow: TCP Ports required to be allowed"
    Write-Host $text -ForegroundColor Yellow
    $text | Add-Content "$Env:TEMP\$fileName_Allow"
    $allowTcpPorts
    $allowTcpPorts | Add-Content "$Env:TEMP\$fileName_Allow"
    Write-Host
    "" | Add-Content "$Env:TEMP\$fileName_Allow"

    $text = "Allow: UDP Ports required to be allowed"
    Write-Host $text -ForegroundColor Yellow
    $text | Add-Content "$Env:TEMP\$fileName_Allow"
    $allowUdpPorts
    $allowUdpPorts | Add-Content "$Env:TEMP\$fileName_Allow"
    Write-Host
    "" | Add-Content "$Env:TEMP\$fileName_Allow"

    Write-Host "-----------------------------------------------------"
    Write-Host

    ####################################################

    $fileName_Default = "o365_Default.txt"

    if(Test-Path "$Env:TEMP\$fileName_Default"){ 
    
        Set-Content -Value "" "$Env:TEMP\$fileName_Default"
        
    }

    $text = "Default: URLs required to be allowed via Proxy"
    Write-Host $text -ForegroundColor Yellow
    $text | Add-Content "$Env:TEMP\$fileName_Default"
    $defaultUrlsCustomObject
    $defaultUrlsCustomObject | Add-Content "$Env:TEMP\$fileName_Default"
    Write-Host
    "" | Add-Content "$Env:TEMP\$fileName_Default"

    $text = "Default: IPv4 Address required to be default"
    Write-Host $text -ForegroundColor Yellow
    $text | Add-Content "$Env:TEMP\$fileName_Default"
    $defaultIpsv4
    $defaultIpsv4 | Add-Content "$Env:TEMP\$fileName_Default"
    Write-Host
    "" | Add-Content "$Env:TEMP\$fileName_Default"

    $text = "Default: IPv6 Address required to be default"
    Write-Host $text -ForegroundColor Yellow
    $text | Add-Content "$Env:TEMP\$fileName_Default"
    $defaultIpsv6
    $defaultIpsv6 | Add-Content "$Env:TEMP\$fileName_Default"
    Write-Host
    "" | Add-Content "$Env:TEMP\$fileName_Default"

    $text = "Default: TCP Ports required to be default"
    Write-Host $text -ForegroundColor Yellow
    $text | Add-Content "$Env:TEMP\$fileName_Default"
    $defaultTcpPorts
    $defaultTcpPorts | Add-Content "$Env:TEMP\$fileName_Default"
    Write-Host
    "" | Add-Content "$Env:TEMP\$fileName_Default"

    $text = "Default: UDP Ports required to be default"
    Write-Host $text -ForegroundColor Yellow
    $text | Add-Content "$Env:TEMP\$fileName_Default"
    $defaultUdpPorts
    $defaultUdpPorts | Add-Content "$Env:TEMP\$fileName_Default"
    Write-Host
    "" | Add-Content "$Env:TEMP\$fileName_Default"

    Write-Host "-----------------------------------------------------"
    Write-Host

    ####################################################

    Write-Host "Optimize output written to '$Env:TEMP\$fileName_Optimize'" -ForegroundColor Gray
    Write-Host "Allow output written to '$Env:TEMP\$fileName_Allow'" -ForegroundColor Gray
    Write-Host "Default output written to '$Env:TEMP\$fileName_Default'" -ForegroundColor Gray
    Write-Host

    #endregion

}

else {

    Write-Host
    Write-Host "Office 365 worldwide commercial service instance endpoints are up-to-date" -ForegroundColor Green
    Write-Host

}
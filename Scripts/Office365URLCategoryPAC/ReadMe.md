# Split Office 365 URL categories to different egress locations 

##### Script Purpose 
The new [WebService](https://docs.microsoft.com/en-us/office365/enterprise/office-365-ip-web-service) gave Microsoft the opportunity to categorize endpoints into levels of importance to allow customers to concentrate on optimization to those endpoints which really require it, thus simplifying the connectivity approach. 
Microsoft strongly recommend core ‘Optimize’ marked traffic is sent direct, not SSL inspected or hairpinned.  These endpoints represent Office 365 scenarios that are the most sensitive to network performance, latency and availability and account for over 80% of the volume of traffic to the service in only a small number of endpoints.
Therefore, a common scenario for a PAC file would be to send these URLs direct, send the ‘Allow’ list to an optimized proxy and the ‘Default’ list to the standard proxy. Alternatively you may want to see the complete list of Allow and Default URLs to perhaps deal with them differently on the same proxy. For example send Allow to port 8081 and don’t SSL inspect, default to port 8080 and treat as normal web traffic. This script allows you to create a PAC file to do just that, programmatically. 


##### Notes

- This script has a number of options when run which are described below and will produce a different output 
- Assumes all Office 365 URLs are required to be handled 
- The script sets the proxy address of 10.10.10.10:8080 in scenario 1 which will require editing to whatever your proxy address is, if used in its output format.
- The Script requires the input of a tenant name to allow for customer specific SharePoint Online/OneDrive for Business URLs to be created rather than *.sharepoint.com (if you don’t input a tenant name *.sharepoint.com is used which covers all tenant names)
- When inputting a tenant name you should only use the unique part. So, if your tenant name is contoso.onmicrosoft.com then you simply enter **contoso** when prompted. 
- If you wish to send more than one tenant’s traffic direct, manually duplicate <yourtenantname>.sharepoint.com & <yourtenantname>-my.sharepoint.com entries and replace yourtenantname with yourothertenantname



##### How do I run the script? 

**Scenario 1:** Create a PAC file to send the ‘Optimize’ marked URLs direct, and all other URLs (Allow & Default) via a single proxy
To obtain a PAC file to achieve this scenario, simply run the script “Office365URLCategorySplitPAC.ps1” in PowerShell and enter the tenant name when prompted. 
**Scenario 2:** Create a PAC file to send the ‘Optimize’ marked URLs direct, ‘Allow’ marked URLs to Proxy A and ‘Default’ marked URLs to Proxy B.
To obtain a PAC file to achieve this scenario, we need to enter some options when running the script.
**Office365URLCategorySplitPAC.ps1 -DefaultProxySettings 10.0.0.1:8080 -SecondaryProxySettings 10.0.0.2:8081 -TenantName contoso -Scenario 2**
Sections in italics are customisable based on your requirements. Default proxy will be for the Allow traffic, secondaryproxy will be for the default marked traffic.

 The script will then output a working PAC file as described above (the file extension will need to be renamed from .txt to .pac first). Output path by default is C:\Users\username\AppData\Local\Temp\



####################################################



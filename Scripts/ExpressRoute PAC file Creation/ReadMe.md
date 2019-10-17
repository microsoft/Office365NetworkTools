
# Send ExpressRoute traffic direct

##### Script Purpose
Customers using Expressroute to connect elements of the Office 365 service (which isn’t the generally recommended connectivity approach for Office 365) will require the ExpressRoute marked URLs to be sent direct so they can traverse the Expressroute link rather than the internet path. This Script, when run, creates a PAC file sending all ExpressRoute supported URLs direct, and all other traffic to a proxy.

##### Notes

- Assumes all Office BGP communities are advertised into the network as the script will output all Expressroutable URLs. Script should be amended if only select BGP communities are used.
-  The script sets the proxy address of 10.10.10.10:8080 which will require editing to whatever your proxy address is if used in its output format.
-  The Script requires the input of a tenant name to allow for customer specific SharePoint Online/OneDrive for Business URLs to be created rather than *.sharepoint.com (if you do not input a tenant name *.sharepoint.com is used which covers all possible tenant names)
-  When inputting a tenant name you should only use the unique part. So, if your tenant name is contoso.onmicrosoft.com then you simply enter **contoso** when prompted.
-  If you wish to send more than one tenant’s traffic direct, manually duplicate <yourtenantname>.sharepoint.com & <yourtenantname>-my.sharepoint.com entries and replace yourtenantname with yourothertenantname
-  The quicktips.skypeforbusiness com URL is required to be sent to a proxy at the top of the PAC before the wildcard *.skypeforbusiness.com is hit as it isn’t Expressroutable whereas the wildcard version is. If you remove this section, the URL will fail to connect as the IP it resolves to has not path via ER. We’re working to remove these inconsistent behaviours in namespaces where possible and this is currently the only one remaining.
-  Every effort is made to ensure 100% accuracy but this script is intended as an example so it is recommended


##### How do I run the script?

Simply run the ps1 file in PowerShell and enter your tenant name when prompted. The script will then output a working PAC file as described above (the file extension will need to be renamed from .txt to .pac first). Output path by default is C:\Users\username\AppData\Local\Temp\

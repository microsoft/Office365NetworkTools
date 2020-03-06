
# Display URL/IP/Port information per endpoint category.

##### Script Purpose
Customers wishing to quickly obtain the URL/IP/Port information per category (Optimize/Allow/Default) to use to seperate the traffic into different routing models. 

##### Notes

- This simple script queries the REST API base web service and outputs the URLs, IPs and ports for each endpoint category, i.e Optimize, Allow and Default. Default endpoints do not have IPs provided so need to be sent either via a proxy or via an unrestricted path. 
- IPV6 addresses arent currently necessary for service operation so can be ignored if not required/desired. 
- The script requests the tenant name as some URLs are tenant specific. If your tenant name were contoso.onmicrosoft.com then input contoso when requested. If you dont put anything wildcards are used for the URL eg *.sharepoint.com instead of contoso.sharepoint.com 
- Optimize marked endpoints should be the priorty to go direct to the service with no inspection. Allow endpoints are required for the service to operate but whilst it's recommended to optimize these, it's not essential. They can be proxied for example. Default endpoints can be treated as normal web traffic.
- Every effort is made to ensure 100% accuracy but this script is intended as an example so it is recommended


##### How do I run the script?

Simply run the ps1 file in PowerShell and enter your tenant name when prompted. The script will then output the information as described above and also create a txt file with the information. The output path by default is C:\Users\username\AppData\Local\Temp\

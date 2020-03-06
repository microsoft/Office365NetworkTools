# /Office 365 Network Bandwidth meter/ 
You can use this solution if you intend to: 
1. Measure network bandwidth usage for pilot users on-boarded to Office 365 
2. Measure network bandwidth usage for services like Exchange Online, SharePoint Online/OneDrive for Business and Microsoft Teams
3. Measure network bandwidth usage for traffic to 'Optimize' category Office 365 endpoints 
4. Measure number of TCP connections used while connecting to Office 365 services like Exchange Online, SharePoint Online/OneDrive for Business
5. Measure number of TCP connections used while connecting to 'Optimize' category Office 365 endpoints 

This solution uses Azure monitoring, specifically Service Map, dependencies for Service Map like Microsoft Monitoring Agent, Dependency Agent are applicable to this solution. This concept assumes you have pilot batch of users on-boarded to Office 365 or you can monitor a subset of user traffic accessing Office 365 services. 

You can apply this concept for measuring any SaaS/PaaS traffic as long as you can filter the traffic based on Process Name or Destination IP endpoints, not limited to Office 365. 

This solution will allow you to monitor and analyse the following example scenarios: 

• Bandwidth used for a particular process or set of processes over a set period of time 

• Bandwidth used by the machine over a set period of time 

• Bandwidth used in connections to a specific port 

• Bandwidth used to a specific IP address or range of addresses 

• IP geolocation of the endpoints connected to 

# Prerequisites 
Azure Subscription

Azure Monitoring/Log analytics workspace  

Microsoft Monitoring Agent (MMA)

Dependency Agent

# New Announcements (December 2019)

1. UDP now supported for measuring network bandwidth usage of Teams media traffic (Audio, Video, Screen Sharing), update your Dependency Agent to the latest version for UDP support
2. Sample Queries for measuring network bandwidth usage and TCP connections to 'Optimize' category Office 365 endpoints

# Support Statement
The scripts, samples, and tools made available through the Open Source initiative are provided as-is. These resources are developed in partnership with the community and do not represent official Microsoft software. As such, support is not available through premier or other official support channels. If you find an issue or have questions please reach out through the issues list and we'll do our best to assist, but there is no support SLA associated with these tools.

# Code of Conduct
This project has adopted the Microsoft Open Source Code of Conduct. For more information see the Code of Conduct FAQ or contact opencode@microsoft.com with any additional questions or comments.



# SharePoint-Online-Continuous-Deployment-Framework
SPO-CDF is a simple provisioning framework for SharePoint Online that represents a lightweight and less demanding alternative to Provider-hosted Add-Ins. 
It supports both ways of deployment, the modern one based on the remote provisioning of customizations and the legacy feature based framework of "no code" sandbox solutions. 

SPO-CDF provides quick path to apply automated branding and customizations to newly created and already existing site collections in the minimal amount of time. 
In addition, it supports postponed creation of site collections from all standard and configurable custom web-templates via a familiar OOB UI (dynamically adjusted).

More detailed descriptions of SPO-CDF can be found in the supplemental document "SharePoint-Online-Continuous-Deployment-Framework.docx".
You can also read these descriptions online in my blog post http://paulborisov.blogspot.fi/2016/10/sharepoint-online-continuous-deployment.html.

TODO list:
- Create simple config file for common parameters and eliminate duplicates in scripts.
- Add seamless support for "on-premises" versions of SharePoint.
- Add support of ClientId / ClientSecret in addition to plain text credentials.
- Add more samples to a demo package (shared Managed Metadata navigation, Managed Properties for Search, local and global Taxonomy).

October 19, 2016: enabled and successfully tested the functionality of SPO-CDF in the "on-premises" environment in addition to "Online" (used the local Farm of SharePoint 2016; it required the Public Update from October 2016 to unify usage of the method Tenant.GetSiteProperties() between "Online" and "on-premises").

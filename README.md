# Project Management Web Part

## Summary

This SharePoint Framework (SPFx) web part provides a robust solution for managing projects directly within SharePoint and Microsoft Teams environments. The web part enables the creation, tracking, and management of project-related information leveraging SharePoint lists. It includes functionality such as advanced filtering, search, and a clean user interface for enhanced user experience.

![alt text](<assets/Screenshot 1.png>)
![alt text](<assets/Screenshot 2.png>)
## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.20.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

- Node.js v16.x or later
- Visual Studio Code
- Gulp CLI
- SharePoint Online tenant

## Solution

| Solution             | Author(s)        |
| -------------------- | ---------------- |
| ProjectManagement    | Andrew Hipwood   |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.0.0   | April 25, 2025   | Initial Release |

## Governance and Deployment

### Deployment Requirements
- All deployments require TAG (Technical Advisory Group) or CAB (Change Advisory Board) review
- External penetration testing is required before deployment
- Site collection app catalog deployment requires additional security review
- Compliance with Group-wide governance processes is mandatory

### Security Considerations
- The solution interacts with SharePoint lists for data storage
- No external service integrations
- Uses standard SPFx libraries and Microsoft Graph API
- All data remains within the Microsoft 365 tenant

### Business Benefits
- Enhanced project management capabilities within SharePoint
- Improved user experience with modern UI components
- Seamless integration with existing SharePoint infrastructure
- Reduced need for external project management tools

### Support and Maintenance
- Regular updates and maintenance will be provided
- Bug fixes and feature enhancements will be managed through the standard CAB process
- Documentation will be maintained and updated as needed

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

1. Clone this repository:
   ```bash
   git clone https://github.com/BIM-Solutions/ProjectManagement.git
   ```
2. Navigate to the solution folder:
   ```bash
   cd ProjectManagement
   ```
3. Run the following commands:
   ```bash
   npm install
   gulp serve
   ```

## Features

This project management web part provides:

- Project listing with search and filtering by project number, name, sector, status, and client
- Integration with SharePoint lists to store and retrieve project data
- Adaptability for both SharePoint pages and Microsoft Teams tabs
- Support for theme variants and responsive design
- Fee tracking and budget management
- Project document management
- Team member assignment and tracking

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft Teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples, and open-source controls for your Microsoft 365 development


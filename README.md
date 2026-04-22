# Word Clause Assistant M365

Microsoft Word Add-in built with Office JS, TypeScript, Express API, Microsoft Graph and SharePoint Online.

## Features
- Load approved clauses from SharePoint
- Insert clauses into Word documents
- Replace placeholders dynamically:
  - {{ClientName}}
  - {{EffectiveDate}}
  - {{GoverningLaw}}

## Architecture
Word Add-in → Express API → Microsoft Graph → SharePoint

## Tech Stack
- Office JavaScript API
- TypeScript
- Node.js / Express
- Microsoft Graph
- SharePoint REST
- GitHub

## Example Use Case
Generate legal clauses from a SharePoint clause library and insert them into contracts with dynamic placeholders.

## Future Enhancements
- AI clause recommendations
- Clause approval workflow
- Dataverse integration
- Adaptive Card notifications

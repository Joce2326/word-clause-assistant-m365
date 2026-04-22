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

## Screenshots

### Word Add-in

<img width="1996" height="1068" alt="image" src="https://github.com/user-attachments/assets/4bdff96c-7f78-46a1-a379-5b06c3dc2291" />

### Placeholder Replacement
<img width="562" height="1156" alt="image" src="https://github.com/user-attachments/assets/838f8d6c-5b53-4e09-abb3-ec4f1ae790fc" />

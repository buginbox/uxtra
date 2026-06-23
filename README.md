# Entra Group Manager SPA

Simple React SPA for managing Microsoft Entra ID group members with a manually pasted Microsoft Graph bearer token.

## Stack

- Node.js 22 via `.nvmrc`
- Vite 8
- React 19
- TypeScript 6

## Run

```bash
nvm use
npm install
npm run dev
```

## What the app does

- Accepts a delegated Graph bearer token pasted by the user
- Loads groups returned by `GET /me/ownedObjects/microsoft.graph.group`
- Updates the selected group name and description from a popup dialog after confirmation
- Loads members for the selected group
- Searches users by first name, last name, display name, mail, or UPN, and also searches security groups for nested group membership
- Adds or removes group members through Microsoft Graph, including security groups inside supported groups

## Expected Graph permissions

Recommended delegated scopes:

- `Group.ReadWrite.All`
- `User.Read.All`

Depending on tenant policy and target group type, Microsoft Graph may still require broader directory permissions or reject the operation.

## Graph calls used by the UI

- `GET /me/ownedObjects/microsoft.graph.group`
- `PATCH /groups/{id}`
- `GET /groups/{id}/members`
- `GET /users?$filter=startswith(...)`
- `GET /groups?$filter=securityEnabled eq true and startswith(...)`
- `POST /groups/{id}/members/$ref`
- `DELETE /groups/{id}/members/{memberId}/$ref`

## Verification

```bash
npm test
npm run lint
npm run build
```

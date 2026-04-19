# MSGraphExcelEditor — Power Platform Custom Connector (PoC)

A proof-of-concept custom connector for fine-grained Excel workbook editing via the Microsoft Graph API.

## Overview

This connector exposes a subset of the Graph API Excel/Workbook endpoints as Power Platform actions. It is intended as a starting point — not a production-ready connector.

## App Registration

An Azure AD app registration is required for the connector's OAuth connection. A sanitized manifest template is included at [`MSGraphExcelConnectorApp(Microsoft Graph format).json`](MSGraphExcelConnectorApp\(Microsoft Graph format\).json).

### Required delegated permissions (Microsoft Graph)

| Permission | Purpose |
| --- | --- |
| `User.Read` | Sign in and read user profile |
| `openid`, `profile`, `email` | Standard OIDC claims |
| `offline_access` | Refresh token support |
| `Files.Read` | Read files the user can access |
| `Files.Read.All` | Read all files the signed-in user can access |
| `Files.ReadWrite` | Read and write files the user can access |
| `Files.ReadWrite.All` | Read and write all files the signed-in user can access |
| `Sites.ReadWrite.All` | Read and write items in all site collections |

### Steps

1. In Azure Portal → **App registrations**, create a new registration (`AzureADMyOrg` audience)
2. Under **Manifest**, paste the contents of the template file — this will set all required permissions in one step
3. Under **Certificates & secrets**, create a client secret and note the value
4. After importing/uploading the custom connector in Power Platform, the platform will generate one or more redirect URIs in the format:

   ```text
   https://global.consent.azure-apim.net/redirect/{connector-name}-{environment-id}
   ```

   Add each generated URI to the app registration under **Authentication → Web → Redirect URIs**
5. Enter the app registration's **Application (client) ID** into `apiProperties.json` at `connectionParameters.token.oAuthSettings.clientId`

> The `id` (object ID), `appId` (client ID), `publisherDomain`, `passwordCredentials`, and `redirectUris` fields in the template are intentionally blank — they are either generated automatically on registration or unique to your environment.

## Getting Started

### Prerequisites

- A registered Azure AD app with `https://graph.microsoft.com/.default` delegated scope
- The app's **Client ID** entered in `apiProperties.json` (`connectionParameters.token.oAuthSettings.clientId`)
- The connector deployed to a Power Platform environment via PAC CLI:

  ```bash
  pac connector create \
    --environment <your-environment-id-or-url> \
    --api-definition-file MSGraphExcelEditor/apiDefinition.Swagger.json \
    --api-properties-file MSGraphExcelEditor/apiProperties.json \
    --icon-file MSGraphExcelEditor/icon.png
  ```

  To add the connector to an existing unmanaged Dataverse solution, append:

  ```bash
    --solution-unique-name <your-solution-unique-name>
  ```

  > `--settings-file` is an alternative but requires `connectorId` and `environment` populated in `settings.json` — use the explicit flags above for a first-time create.

### Typical call sequence

> This connector only works with `.xlsx` files. Legacy `.xls` format is not supported by the Graph API workbook endpoints.

1. **GetSiteByPath** — resolve the SharePoint/OneDrive site and capture the `id` from the response
2. **CreateSession** — open a workbook session; capture the `id` from the response (`workbook-session-id`). Requires `siteId` and the relative path to the `.xlsx` file.
3. Call any read/write operations, passing `siteId`, the relative path to the `.xlsx` file, and `workbook-session-id` on each request
4. **CloseSession** — close the session when done to commit or discard changes. Requires `siteId` and the relative path to the `.xlsx` file.

## Design Decisions

### Why `sites('{siteId}')` instead of `/drives/{driveId}`

The Graph API exposes workbook operations under several base paths:

| Base path | Works for workbook operations? |
| --- | --- |
| `/v1.0/drives/{driveId}/...` | No — returns 403 "Could not obtain a WAC access token" |
| `/v1.0/users/{userId}/drive/...` | Yes |
| `/v1.0/me/drive/...` | Yes |
| `/v1.0/groups/{groupId}/drive/...` | Yes |
| `/v1.0/sites('{siteId}')/drive/...` | Yes — works universally |

All of the working paths (`users`, `me`, `groups`, `sites`) can be resolved to a site ID via **GetSiteByPath**, so using `siteId` as the single addressing scheme avoids duplicating operations across multiple base paths.

### Resolving the Site ID

Use **GetSiteByPath** first. It takes a `hostName` (e.g. `contoso.sharepoint.com`) and a `relativePathFromHost` (e.g. `sites/MyTeamSite`) and returns the site object. The `id` property in that response is the `siteId` required by all subsequent workbook operations.

### Workbook Sessions

The Graph API supports persistent workbook sessions that batch changes before committing them. Creating a session with `persistChanges: true` means edits are written to the file. Pass the session `id` as the `workbook-session-id` header on each subsequent request to keep operations within the same session context, then call **CloseSession** to finalize.

## Adding New Operations

### URL pattern

Every workbook operation follows this structure:

```text
/v1.0/sites('{siteId}')/drive/root:/{pathFromRoot}:/workbook/{resource}
```

`pathFromRoot` is the file path relative to the drive root, e.g. `Documents/Budget.xlsx` (no leading `/`). This is preferred over addressing by `itemId` because the path is typically already known to the caller — resolving an `itemId` would require an additional lookup operation. The workbook sub-resources nest under it:

```text
/workbook
  /worksheets/{idOrName}
    /cell(row={row},column={column})       ← zero-based index
    /range(address='{A1notation}')         ← e.g. 'A1:C5' or 'Sheet1!A1:C5'
    /usedRange                             ← bounding box of all non-empty cells
    /tables/{idOrName}
      /rows
      /columns/{idOrName}
  /tables/{idOrName}
  /names/{name}                            ← named range or named item
  /functions                               ← call worksheet functions
```

Both IDs and names are accepted wherever `{idOrName}` appears. Worksheet and chart IDs contain curly braces and must be URL-encoded when used in a path.

### Addressing ranges

The Graph API offers several ways to address a cell or region. Choose based on what information the caller is likely to have:

| Method | Swagger path segment | Notes |
| --- | --- | --- |
| A1 notation | `/range(address='{address}')` | Most familiar; supports cross-sheet refs (`Sheet1!A1:B2`) |
| Row/column index | `/cell(row={row},column={column})` | Zero-based; returns a single-cell range object |
| Named range | `/names/{name}/range` | Caller needs to know the defined name |
| Used range | `/usedRange` | No parameters; good for reading dynamic data |
| Table rows/columns | `/tables/{id}/rows`, `/columns/{id}` | Structured access; append a row by POSTing with `index: null` |

**Avoid unbounded ranges** (e.g. `C:C`, `1:4`) in write operations — the API rejects them. They are valid for reads but return `null` for cell-level properties.

### Read vs. write

- **GET** operations return the full range object (`values`, `text`, `formulas`, `numberFormat`, etc.)
- **PATCH** updates a range; passing `null` in the values array leaves that cell unchanged; passing `""` clears it
- **POST** is used for actions (createSession, closeSession, function calls) and for appending rows to tables
- Always pass `workbook-session-id` as an optional header on every operation — even reads benefit from session context

### Body schema: scalars vs. arrays

The Graph API technically requires `values`, `formulas`, and `numberFormat` as 2D arrays (an array of rows, each row an array of cell values). However, in practice:

- **Single-cell operations** (`/cell(row,column)`) — the API accepts plain scalar strings. Define body properties as `type: string` in the Swagger. This produces a much friendlier connector experience in Power Automate (simple text fields rather than expression-built arrays).
- **Multi-cell / row operations** (`/range(address=...)`, table rows, etc.) — scalar coercion is unlikely to hold. Use the proper 2D array schema (`array` of `array`) and instruct callers to construct the array expression in their flow.

When in doubt, test with a scalar first. If the API rejects it, fall back to the 2D array schema.

#### Policy templates as an alternative

If an operation requires a body shape that is impractical to construct in a Power Automate expression (e.g. wrapping a scalar into `[[value]]`), Power Platform [custom connector policy templates](https://learn.microsoft.com/en-us/connectors/custom-connectors/policy-templates) can transform the request body between what the connector exposes to the caller and what gets sent to the API. This lets the Swagger define a user-friendly flat schema while the policy handles the structural transformation transparently. Three built-in templates are relevant to body shape transformations in this connector:

| Template | Direction | What it does |
| --- | --- | --- |
| [`convertobjecttoarray`](https://learn.microsoft.com/en-us/connectors/custom-connectors/policy-templates/convertobjecttoarray/convertobjecttoarray) | Request or Response | Converts a JSON object (key→value pairs) into an array of objects |
| [`convertarraytoobject`](https://learn.microsoft.com/en-us/connectors/custom-connectors/policy-templates/convertarraytoobject/convertarraytoobject) | Request or Response | Converts an array into a JSON object keyed by a chosen property — inverse of the above |
| [`stringtoarray`](https://learn.microsoft.com/en-us/connectors/custom-connectors/policy-templates/stringtoarray/stringtoarray) | Request or Response | Splits a delimited string into an array of objects — useful if a caller passes comma-separated values that the API expects as an array |
| [`setproperty`](https://learn.microsoft.com/en-us/connectors/custom-connectors/policy-templates/setproperty/setproperty) | Request or Response | Adds or overwrites a property on a JSON object with a constant, a value copied from another property, or a template combining both — useful for injecting a wrapper property around a scalar before it reaches the API |

For the `UpdateCell` use case (wrapping a scalar into `[[value]]`), none of these are a direct fit — they operate on existing objects/arrays rather than constructing new nesting. If the scalar coercion stops working, the most pragmatic fallback is to define the body property as a string, document that callers must wrap it (e.g. `[[triggerBody()?['value']]]`), and reference the policy templates as a place to look for a cleaner solution.

### Adding an operation to the Swagger

1. Copy an existing path block of the same HTTP method as a starting point
2. Replace the path, `operationId`, `summary`, and `description`
3. Keep `siteId` and `pathFromRoot` as required path parameters — every workbook operation needs them
4. Add `workbook-session-id` as an optional (or required, for CloseSession) header parameter
5. Define the response schema from the Graph API docs; an empty `schema: {}` works for prototyping
6. For write operation bodies on single-cell endpoints, define properties as `type: string` with a friendly `title`; for multi-cell endpoints use nested arrays

## Current Operations

| Operation | Method | Description |
| --- | --- | --- |
| `GetSiteByPath` | GET | Resolve a site by hostname and path |
| `CreateSession` | POST | Open a workbook editing session |
| `CloseSession` | POST | Close and commit/discard a session |
| `ListWorksheets` | GET | List all worksheets in a workbook |
| `ListNames` | GET | List named ranges and named items |
| `ListTables` | GET | List tables in a workbook |
| `GetCell` | GET | Get a cell by zero-based row/column index |
| `UpdateCell` | PATCH | Update a cell's value, formula, and/or number format |
| `GetRange` | GET | Get a range by A1 notation (values, formulas, text, numberFormat) |
| `UpdateRange` | PATCH | Update a range by A1 notation — values, formulas, and/or number format as 2D arrays |
| `CreateWorksheet` | POST | Add a new worksheet to the workbook |
| `UpdateWorksheet` | PATCH | Rename a worksheet or change its visibility |
| `CreateTable` | POST | Create a table from an A1-notation range |
| `UpdateTable` | PATCH | Rename a table or toggle headers/totals/style |
| `GetNamedRange` | GET | Get a single named range by name |
| `CreateNamedRange` | POST | Define a new named range with a reference and optional comment |
| `UpdateNamedRange` | PATCH | Update a named range's reference, visibility, or comment |

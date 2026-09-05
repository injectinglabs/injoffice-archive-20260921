# XLSX connector extension v1

InjOffice persists its credential-free `ConnectorSpec` subset in a custom OPC
part because the project does not claim that its HTTP/JSON/CSV gateway model is
equivalent to Excel external connections or query tables.

The v1 package identities are fixed:

| Item | Value |
|---|---|
| Part | `/customXml/injofficeConnectors.xml` |
| Workbook relationship | `https://schemas.injoffice.dev/relationships/connectors/2026` |
| Content type | `application/vnd.injoffice.connectors+xml` |
| XML namespace | `https://schemas.injoffice.dev/xlsx/connectors/2026` |

The workbook owns exactly one internal relationship to exactly one connector
part. The part has an explicit content-type override. Its root is
`connectors`, with `version="1"` and `encoding="base64-json"`; the text is a
strict base64 encoding of this JSON envelope:

```json
{"version":1,"connectors":[]}
```

The JSON connector records use the public `ConnectorSpec` field names. Unknown
XML, attributes, JSON fields, versions, relationship targets, duplicate IDs,
or multiple owners are refused instead of overwritten. Definitions are bounded
to 1,024 items and native sheet coordinates. Only credential-free HTTP/HTTPS or
host-relative gateway URLs without userinfo, query parameters, fragments, or
backslashes are accepted.

`SetConnectorDefinitions` changes only the owned part and, when creating or
removing it, its exact workbook relationship and content-type declaration. The
generic `xlsxpatch.Apply` path raw-copies and verifies every untouched ZIP part.
Updating definitions replaces only the connector part. Removing definitions
does not remove any unknown custom XML part or relationship.

Fetched cell data, access tokens, cookies, request headers, authorization state,
errors, cache contents, and last-refresh status are never stored. Hosts must keep
credentials and egress policy in their injected fetcher. The extension provides
InjOffice round-trip persistence; it is not represented as native Excel Power
Query, Data Model, connection, or query-table interoperability.

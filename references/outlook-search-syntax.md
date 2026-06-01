# Outlook Search Syntax — Microsoft Graph

Reference for `$search` and `$filter` query parameters on the
`GET /me/messages` endpoint.

## KQL `$search` Syntax

The `$search` parameter uses Keyword Query Language (KQL). Wrap the
entire value in double quotes in the URL query string.

### Fields

| Field | Example |
|-------|---------|
| `from:` | `from:alice@example.com` |
| `to:` | `to:bob@example.com` |
| `cc:` | `cc:team@example.com` |
| `bcc:` | `bcc:secret@example.com` |
| `subject:` | `subject:quarterly report` |
| `body:` | `body:action items` |
| `participants:` | `participants:alice` (any from/to/cc) |
| `attachment:` | `attachment:budget.xlsx` |
| `hasAttachments:true` | Messages with attachments |
| `sent:` | `sent:2026-01-15` or `sent>=2026-01-01` |
| `received:` | `received>=2026-05-01` |

### Operators

- **AND** (implicit): `from:alice subject:report` → both must match
- **OR**: `from:alice OR from:bob`
- **NOT**: `NOT from:noreply@example.com`
- **Phrases**: `subject:"Q4 review"` (use double quotes)

### Examples

```
$search="from:boss@company.com subject:urgent"
$search="hasAttachments:true received>=2026-05-01"
$search="subject:\"weekly standup\" NOT from:noreply"
$search="participants:alice body:\"action items\""
```

## OData `$filter` Syntax

The `$filter` parameter uses standard OData operators. Can be combined
with `$search` in some cases, but Graph may reject complex combos.

### Common Filters

| Filter | Example |
|--------|---------|
| Unread | `isRead eq false` |
| Has attachments | `hasAttachments eq true` |
| Flagged | `flag/flagStatus eq 'flagged'` |
| Importance | `importance eq 'high'` |
| Date range | `receivedDateTime ge 2026-05-01T00:00:00Z` |
| Sender | `from/emailAddress/address eq 'alice@example.com'` |
| Subject contains | `contains(subject, 'report')` |

### Operators

| Operator | Meaning |
|----------|---------|
| `eq` | Equal |
| `ne` | Not equal |
| `gt` | Greater than |
| `ge` | Greater than or equal |
| `lt` | Less than |
| `le` | Less than or equal |
| `and` | Logical AND |
| `or` | Logical OR |
| `not` | Logical NOT |
| `contains()` | Substring match |
| `startsWith()` | Prefix match |

### Examples

```
$filter=isRead eq false and importance eq 'high'
$filter=receivedDateTime ge 2026-05-01T00:00:00Z and receivedDateTime lt 2026-06-01T00:00:00Z
$filter=hasAttachments eq true and from/emailAddress/address eq 'alice@example.com'
```

## `$orderby` and `$select`

Commonly combined with search/filter:

```
$orderby=receivedDateTime desc
$select=id,subject,from,receivedDateTime,isRead,bodyPreview
$top=25
```

## Pagination

Graph returns `@odata.nextLink` when more results are available.
Follow this URL to get the next page.

## Notes

- `$search` is best for keyword/free-text queries.
- `$filter` is best for structured/boolean queries (unread, date range).
- Combining both in one request can fail — prefer one or the other.
- Maximum `$top` is 1000, but Graph may return fewer.
- Default sort for search results is by relevance, not date.

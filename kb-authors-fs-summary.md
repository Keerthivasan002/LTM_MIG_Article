# Freshservice article and author lookup

Lookup of live Freshservice articles and current authors for:

`kb_knowledge Authors and Suggested Authors for Missing Ones.xlsx`

Matched by ServiceNow **Sys ID** (1174) or **KB number** (7) against `New folder/migration-results-live.csv`, then `GET /api/v2/solutions/articles/{id}` on **amexgbt.freshservice.com**.

## Result

| | Count |
|---|---|
| Rows in the workbook | 1181 |
| Freshservice articles found | 1181 |
| Missing article | 0 |
| Missing author | 0 |
| Unique Freshservice authors | 15 |

Original ServiceNow **Author** / **Email** columns were left unchanged.

## New columns

| Column | Description |
|---|---|
| FS Article ID | Freshservice solutions article id |
| FS Article URL | Agent article URL |
| FS Article Title | Title currently in Freshservice |
| FS Author | Current article author name in Freshservice |
| FS Author Email | Current author email |

## Current Freshservice authors

| FS Author | Articles |
|---|---|
| Rohini Thombre | 803 |
| Danilo Mazzolani | 188 |
| Ernielyn Villanueva | 150 |
| Michael Galvin | 16 |
| Matthew Mueller | 6 |
| Ian Edwards | 5 |
| Heather Mcfadden | 3 |
| Daniel Walsh | 2 |
| April McGuire | 2 |
| Ben Ptacek | 1 |
| Others | 5 |

## Example

KB51341 still has ServiceNow author Alexis Salazar (`ASalazar@mycwt.com`). Freshservice article [17000209600](https://amexgbt.freshservice.com/a/solutions/articles/17000209600) (*Nitro Monitoring Incident- (Malware_Persistence)*) is authored by **Michael Galvin** (`michael.galvin@amexgbt.com`).

Re-run with:

```powershell
node fill-fs-article-authors.js
```

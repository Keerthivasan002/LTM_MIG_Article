# Chris group list vs live Freshservice

Compared **Group Members List V3 for Chris Peterson.xlsx** with live Freshservice on **amexgbt.freshservice.com** (agent workspaces, members, and observers).

Chris has **4534** member rows across **843** groups. Groups were matched with or without a `CWT` prefix (same as Hari’s list). All matched groups are in workspace **My Team**.

## Result

| | Count |
|---|---|
| Chris groups matched in FS | 734 |
| Chris groups with no FS group | 109 |
| Chris rows that are FS members | 3802 |
| Chris rows that are FS observers only | 0 |
| Chris rows not in that FS group | 81 |
| Chris emails not found as FS agents | 445 |
| Chris rows whose group did not match | 206 |
| FS members of matched groups not in Chris | 919 |
| FS observers on matched groups | 0 |

Chris’s file is members-only (no observer column). Live FS also has no observers on these matched groups.

## Unmatched Chris groups

Unmatched groups include the AQUA / Genpact-Approvers / TX Op Error sets, plus several already-prefixed CWT names that still do not exist in live, for example:

- CWT Digital Leadership Approval Group
- CWT FP&A MAINTENANCE

## Files

- `chris-vs-fs-groups.xlsx` — Summary, Group match, Unmatched Chris groups, Chris members vs FS, FS members not in Chris, FS observers
- `chris-vs-fs-groups-members.csv`

Re-run with:

```powershell
node compare-chris-fs-groups.js
```

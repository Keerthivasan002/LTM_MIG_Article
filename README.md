# LTM_MIG_Article

Migrate ServiceNow knowledge articles into Freshservice, including attachments and inline images.

## Setup

1. Copy `.env.example` to `.env` and fill in ServiceNow and Freshservice credentials.
2. Set `FS_FOLDER_ID` to an existing Freshservice solutions folder.
3. Put article `sys_id` values in `articles.xlsx` under the `SNOW_ARTICLE_SYS_ID` column.
4. Install dependencies and run:

```powershell
npm install
node migrate-sample.js
```

`.env` and `node_modules` are not committed. Successful rows are skipped on later runs; see `migration-results.csv` on your machine.

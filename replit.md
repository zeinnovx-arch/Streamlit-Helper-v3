# Excel Filter App

An interactive Streamlit dashboard for uploading Excel workbooks or CSV files, filtering rows, reviewing summaries, and exporting results.

## Run & Operate

- `streamlit run app.py --server.port 5000 --server.address 0.0.0.0` — run the Excel filtering dashboard
- `pnpm --filter @workspace/api-server run dev` — run the API server (port 5000)
- `pnpm run typecheck` — full typecheck across all packages
- `pnpm run build` — typecheck + build all packages
- `pnpm --filter @workspace/api-spec run codegen` — regenerate API hooks and Zod schemas from the OpenAPI spec
- `pnpm --filter @workspace/db run push` — push DB schema changes (dev only)
- Required env: `DATABASE_URL` — Postgres connection string

## Stack

- Streamlit, Python 3.11, pandas, openpyxl, xlrd
- pnpm workspaces, Node.js 24, TypeScript 5.9
- API: Express 5
- DB: PostgreSQL + Drizzle ORM
- Validation: Zod (`zod/v4`), `drizzle-zod`
- API codegen: Orval (from OpenAPI spec)
- Build: esbuild (CJS bundle)

## Where things live

- `app.py` — Streamlit UI, workbook loading, filtering, summaries, and exports
- `pyproject.toml` — Python dependencies
- `main.py` — original Python scaffold entry point; not used by the dashboard workflow

## Architecture decisions

- Uploaded files are processed in memory for the current session and are not persisted.
- Filtering keeps missing values visible when a numeric, date, or categorical filter is active.
- Excel exports are generated in memory as a new workbook containing only the filtered rows.

## Product

- Upload `.xlsx`, `.xls`, or `.csv` files.
- Switch between workbook sheets.
- Search all columns and add targeted text, numeric-range, date-range, or categorical filters.
- View row counts, match rate, column metadata, numeric summaries, and the filtered table.
- Download filtered results as CSV or Excel.

## User preferences

No additional preferences recorded.

## Gotchas

- The dashboard workflow uses Streamlit on port 5000 with a public bind address.
- Legacy `.xls` files use the `xlrd` engine; `.xlsx` files use `openpyxl`.

## Pointers

- See the `pnpm-workspace` skill for workspace structure, TypeScript setup, and package details

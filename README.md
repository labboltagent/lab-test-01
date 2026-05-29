# lab-test-01

Minimal Excel (3-tab) → DB importer.

## Excel format (minimal)

- 3 sheets (tabs): first is parent, second child, third grandchild.
- Each sheet must have an `id` column.
- Parent sheet has a `child_ids` column with comma-separated child ids.
- Child sheet has a `grandchild_ids` column with comma-separated grandchild ids.
- The importer writes an `__status__` column per row (`inserted` or `error: ...`).

## DB expectations (minimal)

- Tables exist with names matching the sheet names.
- Join tables exist:
  - `{parent}_{child}` with columns `{parent}_id`, `{child}_id`
  - `{child}_{grandchild}` with columns `{child}_id`, `{grandchild}_id`

## Run

Run importer (SQLite example):

`python3 -m excel_hierarchy_importer path/to/file.xlsx --sqlite path/to/db.sqlite`

Run unit tests:

`python3 -m unittest discover -s tests -v`

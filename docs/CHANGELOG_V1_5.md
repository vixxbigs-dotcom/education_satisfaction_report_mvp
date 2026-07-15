# v1.5 PPT table height fix

## Fixes

- Fixed downloaded PPT tables stretching vertically when the table has only a few rows.
- Curriculum and survey-structure tables now use dynamic table heights based on row count.
- Explicit row heights are applied after table creation so PowerPoint does not distribute a small number of rows across the full slide height.

## Files changed

- `src/ppt_renderer.py`
- `docs/CHANGELOG_V1_5.md`

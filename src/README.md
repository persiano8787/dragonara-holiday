# Sources

- `core.js`: calendar reader shared by the manager app (rules in the header comment).
- `manager_template.html` + `core.js` + `logo.txt` -> `manager.html` (replace `/*__CORE__*/` and `__LOGO__`).
- `gen_data.js`: rebuilds `data/*.json` from a folder of Excel files (Node + jszip + @xmldom/xmldom).
- `test_core.js`: compares the reader with the reference results.
- `staff_template.html`: design source of the staff page.

No Excel files or names are stored in this repository.

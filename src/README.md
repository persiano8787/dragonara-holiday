# Sources

- `core.js`: calendar reader shared by the manager app (rules in the header comment).
- `edit.js`: surgical .xlsx editor (cell fill, threaded comments with legacy note and VML shape); everything else stays byte-for-byte.
- `manager_template.html` + `core.js` + `edit.js` + `logo.txt` -> `manager.html` (replace `/*__CORE__*/`, `/*__EDIT__*/` and `__LOGO__`).
- `gen_data.js`: rebuilds `data/*.json` from a folder of Excel files (Node + jszip + @xmldom/xmldom).
- `test_core.js`: compares the reader with the reference results.
- `test_publish.js`: Save & publish in real Chromium with a browser folder and a mocked GitHub API (checks that no names are published).
- `test_edit.js`, `test_edit2.js`: edit copies of the calendars; `test_ui2.js`: full app run in jsdom with a fake folder (save, backup, lock and stale checks).
- `staff_template.html`: design source of the staff page.

No Excel files or names are stored in this repository.

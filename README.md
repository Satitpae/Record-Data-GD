# GD Record

Google Apps Script project for the **GD Record** field-operations system.

## Files

- `Code.gs` — server-side Google Apps Script. Reads/writes Google Sheets, validates records, handles create/update operations, and provides audit helpers.
- `Index.html` — user interface for reports, plans, customers, and customer contacts.

## Use in Google Apps Script

1. Open the Apps Script project bound to the target Google Sheet.
2. Replace the contents of the script file with `Code.gs`.
3. Create an HTML file named `Index` and paste in `Index.html`.
4. Save the project and reload the spreadsheet to use the **GD Record** menu.
5. If deploying as a Web App, deploy a new version from Apps Script after updating the files.

## Repository note

`Code.gs` currently contains the Spreadsheet ID used by the project. If this repository will be public, consider moving environment-specific values to Apps Script Properties before publishing.

## Source conversion

These files were converted from Markdown-pasted source. Markdown-only escaping was removed while JavaScript regular-expression escapes such as `\d` and `\s` were preserved.

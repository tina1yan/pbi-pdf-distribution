# Public-package privacy review

Scope: the eight allowlisted text files in this release. The original source and local sanitization script are outside the release directory and must never be added to a Git commit or archive.

| Item | Public release treatment |
|---|---|
| Department names and codes | Replaced with synthetic bookmarks A, B, C and D |
| Report page labels | Replaced with Page A/B/C and Volume A/B/C |
| Workbook/project label | Replaced with ABCD or generic PBI labels |
| Filesystem examples | Only `C:\abcd\input` and `C:\abcd\output` |
| Actual PBIX and Excel workbooks | Excluded |
| Business data, PDFs and screenshots | Excluded |
| Original UI diagnostics/model labels | Excluded |
| Full runtime UI snapshots | Disabled by default in the public source |
| Source-to-example replacement mapping | Kept out of the public package |
| Credentials and personal contact details | None included |

The source reads operational paths and filenames from a private Excel control sheet. It uses Windows' TEMP environment variable for the Power BI print-job directory; this is a generic runtime path, not a published username or machine location.

## Checks

Run `python tools/check_release.py` before publishing. It verifies the file allowlist, rejects UNC paths, non-example absolute Windows paths, email addresses and selected credential formats, and checks source structure and key public settings. Synthetic negative controls check that the scanner detects representative leaks. The scanner does not replace manual review or guarantee that arbitrary future business information will be caught.

The prepared package also received a local comparison against the recovered source for original report identifiers and workbook labels. No raw original values are reproduced in this review or in the public checker.

## Runtime privacy boundary

The repository is a sanitized example; running it against a real report will create real outputs and may record real paths or labels in the private workbook. Hidden worksheets are not anonymization or access control. `.gitignore` reduces accidental additions but does not scrub files, remove tracked files or cleanse Git history.

Publish only the allowlisted source/docs from a clean commit. Do not publish an existing workbook, even if its visible sheet names were changed. Workbook metadata, hidden sheets, connections and VBA strings can still contain private information.

## Validation limits

The source was inspected statically. No Windows Excel/Power BI/Acrobat runtime was available for compilation or end-to-end testing of this public adaptation. The new example setup module is also untested in Excel. The original Adobe splitting module has not been recovered, so it is not included or marked as reviewed.

# Method 2 setup

## Requirements

- Windows with desktop Excel supporting VBA7 (`PtrSafe` and `LongPtr`).
- Power BI Desktop using English UI labels compatible with this module.
- In Excel's VBA editor: Tools > References > select the UIAutomationClient / Microsoft UI Automation type library. The code uses early-bound `UIAutomationClient` types.
- A local, private test PBIX and an empty output folder. No example PBIX is supplied.

Import `src/PBI_UI_Export.bas` and `src/PBI_Example_Control.bas` via VBA File > Import File. If pasting into a standard module instead, omit its `Attribute VB_Name` line. Save the workbook as `.xlsm` in your own private folder. Compile the VBA project before running.

## Synthetic report contract

Use a private report copy. Create the following exact page names:

- `Page A`, `Page B`, `Page C` — base pages included in every export.
- `Volume A`, `Volume B`, `Volume C` — optional pages for the matching bookmark.

Configure bookmarks as follows. Each bookmark must have the correct filters/date state saved in Power BI; the macro does not create or update bookmarks.

| Bookmark | Optional page | Expected exported pages |
|---|---|---:|
| A | Volume A | 4 |
| B | Volume B | 4 |
| C | Volume C | 4 |
| D | blank | 3 |

Other discovered pages are hidden for export. The macro identifies hidden page labels through the icon prefix exposed by Power BI's accessibility tree. It must be able to discover and read all page tabs; it does not require that every page initially be unhidden. It stops if a state is missing or ambiguous. This also excludes extra legacy/helper pages without hard-coding their business names.

The total number of discovered template pages must be at most 50. The UI approach repeats this small template for each bookmark; it cannot export an arbitrary single report larger than Power BI's supported limit.

## Control sheet

Run `PBI_CreateExampleControl` in a new workbook. It creates `ABCD Control` without changing an existing sheet. It sets all Run flags to `No` and does not create folders or reports.

| Cell | Example | Meaning |
|---|---|---|
| B16 | `C:\abcd\input` | PBIX folder |
| B17 | `ABCD.pbix` | PBIX file name |
| A19 | Filename | PDF filename column |
| B19 | Folder | Existing output directory column |
| C19 | Bookmark | Exact bookmark name column |
| D19 | Volume Tab | Exact optional page name column |
| E19 | Run | `Yes` selects a row |
| F19 | Status | Macro writes progress/result here |

Rows 20–23 contain synthetic examples with `C:\abcd\output` as destination. Replace these paths only in your private workbook. Keep any date-based filename formula in column A; the macro uses its evaluated text and does not calculate a reporting date itself.

1. Create the input and output folders and supply your test PBIX.
2. Open that exact PBIX in Power BI Desktop. Ensure the report page bar is present and bookmark groups are expanded.
3. Refresh and check the date/filters manually. Configure each bookmark correctly.
4. In Excel, set row 20's Run value to `Yes`, select a cell in that row and run `PBI_Test_Selected_Department`.
5. Leave the desktop alone while the macro works. Hold Escape to request a stop.
6. Open the PDF and verify department/bookmark, dates, figures, page count and visual rendering. “PDF created - REVIEW” is deliberate, not a certification of content.
7. Set other desired rows to `Yes` and run `PBI_Batch_Export_Departments`.

## File handling and recovery

The macro watches `%TEMP%\Power BI Desktop\print-job-*` for a new job with the expected report PDF name. This is implementation-dependent desktop behavior, not a supported Power BI export API. The folder must already exist. A fresh installation or future version may require a successful manual export first or code adaptation.

The file must stop changing for 10 seconds and pass basic PDF header/end-marker checks. The macro then stages a copy, compares its bytes and installs it at the destination. There is a 240-second export timeout. These checks detect common incomplete-copy cases; they do not parse or validate all PDF structures.

Existing destination files are replaced using a temporary backup. The backup is removed after a successful replacement. If a `.pbipartial` file remains, inspect or move it before retrying. The output from previously completed rows is not rolled back on a later failure.

Failures retain a hidden `PBI_UI_...` worksheet. Unhide it locally to diagnose the issue. Logs can contain full paths and UI labels; never publish them or your working workbook without a new privacy review. Full UI snapshots are disabled by `PBI_VerboseSnapshot = False`; enabling them can collect model and report labels.

After a failure, inspect the report state before retrying. The macro has not saved the PBIX, but the open report view may have changed.

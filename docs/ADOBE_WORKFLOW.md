# Method 1: export once, split with Adobe Acrobat

**Original VBA source pending.** This document records the requested workflow, not a recovered implementation. No substitute macro is represented as the original tested version.

1. Prepare the report and export a combined PDF from Power BI.
2. Keep that export below 50 pages under this workflow's operating rule.
3. Use Excel VBA with the Adobe Acrobat automation library to open the PDF and split configured page ranges.
4. Name each output with its generic report identifier and reporting date, and save it to its configured destination folder.
5. Verify the split page ranges and content before distributing the PDFs.

Example identifiers are A, B, C and D. Example locations are `C:\abcd\input` and `C:\abcd\output`. The actual split ranges must come from the report's validated page order, not invented example numbers.

The original module is needed to confirm its exact Acrobat reference, object calls, workbook layout, overwrite behavior and error handling. Adobe documents `AcroExch.PDDoc` and operations including `InsertPages` and `Save`, but that does not establish which calls the original module used. Do not assume the free Reader supplies the same automation capabilities as Acrobat.

The page-count constraint applies to the upstream Power BI export. Adobe's splitting library should not be described as having a general 50-page limit. Microsoft's documented condition is **more than 50 report pages**; this workflow uses the stricter **fewer than 50** convention.

Sources: [Microsoft](https://learn.microsoft.com/en-us/power-bi/collaborate-share/end-user-pdf), [Adobe](https://opensource.adobe.com/dc-acrobat-sdk-docs/library/interapp/IAC_API_OLE_Objects.html).

# PBI to PDF Distribution

Two approaches to preparing Power BI PDFs for distribution: export a combined PDF and split it with Adobe Acrobat, or replay the desktop workflow with Windows UI Automation.

**Release status:** the sanitized UI Automation module is included. The original Adobe-splitting VBA module has not yet been recovered; method 1 is documented only. This is not a claim that both implementations are available or tested in this release.

| | Method 1: Export + Adobe split | Method 2: UI Automation |
|---|---|---|
| Workflow | Export a combined PDF, then use Excel VBA and the Adobe Acrobat automation library to split and name the outputs | Use Excel VBA and Windows UI Automation to select bookmarks, adjust page visibility, export and copy each PDF |
| Best fit | A stable page order and a combined export small enough for Power BI | Reusing a compact report template across many bookmarks |
| Page limit | This workflow uses **fewer than 50 pages** as its operating rule | The recovered module stops if it discovers more than 50 template pages; it does not bypass Power BI export limits |
| Model tokens at runtime | None for a local VBA/Acrobat workflow | **0 model tokens**: no LLM, model API, prompt loop or API key is used by the included macro |
| Desktop dependency | Acrobat automation must be available; the original module's exact requirements still need verification | Power BI must remain open in an active, unlocked Windows desktop session |
| Included here | [Workflow notes](docs/ADOBE_WORKFLOW.md); original source pending | [Sanitized VBA source](src/PBI_UI_Export.bas) and [setup guide](docs/SETUP.md) |

## An agent-like workflow with zero runtime model tokens

Method 2 is deterministic robotic process automation (RPA). It imitates a user's clicks using Windows UI Automation to find controls, Win32 calls to focus the window and move the mouse, and keyboard input where needed. It does not reason about the report or learn a changed interface.

It finds the already-open report, maximizes it, prepares the Bookmarks pane, applies each selected bookmark, exposes three base pages plus an optional volume page, exports a PDF, waits for a new completed print job, and copies the result to the configured destination. It checks PDF header/end markers and compares copied bytes. **It does not independently validate the exported PDF's page count, filters, dates, figures or recipient suitability.**

Zero tokens refers to the macro's execution, not the tools used to develop it, software licensing, computer resources, or the Power BI report's own data connections.

## Important operating limits

- Keep the intended Power BI report open. The macro does not launch Power BI or refresh the data.
- Keep Windows unlocked and the interactive desktop available. Do not switch applications or use the mouse/keyboard during the run. Minimized or disconnected remote sessions are not supported by this release.
- Use one Power BI export at a time on the Windows account. The macro rejects competing new print jobs.
- The selectors use English UI text and particular control IDs. Power BI updates, pane layouts, display scaling, missing tabs, collapsed bookmark groups or language changes can break them.
- Bookmark selection sends input; it cannot prove that the bookmark saved the intended filters. Review a single output before a batch.
- It changes the open report's page visibility and bookmark state. It does not save the PBIX or restore its initial view.
- Successful runs delete their diagnostic worksheet. Failures keep it hidden. **Hidden is not anonymized:** runtime logs and the control worksheet may contain local paths and report labels. Keep working files private.
- Existing destination PDFs can be replaced. Run the example in an empty test folder first. Earlier successful outputs remain if a later row fails.
- This release prepares PDFs in folders. It does not email them or select recipients.

## Quick start

1. Read [SETUP.md](docs/SETUP.md).
2. Import both `.bas` modules into a new, empty macro-enabled Excel workbook.
3. Add the **UIAutomationClient / Microsoft UI Automation** reference in the VBA editor.
4. Run `PBI_CreateExampleControl` to create a synthetic control sheet.
5. Use your own private report copy with the example names, check the configured paths, then test one row with `PBI_Test_Selected_Department`.
6. After checking its output, run `PBI_Batch_Export_Departments` for rows marked `Yes`.

No PBIX, workbook, PDF, screenshot, source log, real department mapping or private filesystem path is included. All report examples use A/B/C/D and `C:\abcd\...`.

## Validation and provenance

The UI module was adapted from an existing version 1.14 source. Public changes are generic names, a four-bookmark example mapping, generic diagnostic names, and verbose UI snapshots disabled by default. The example control-sheet builder is new. The original working files were not modified.

Only static source/privacy checks have been performed on this public adaptation. **It has not been compiled or run against Windows, Excel, Power BI or Acrobat in this preparation environment.** See [privacy review](docs/PRIVACY_REVIEW.md). The Adobe module is still pending; no reconstructed replacement is presented as the original.

Microsoft documents that reports with **more than 50 report pages** cannot be exported to PDF. The stricter “fewer than 50” rule above is this workflow's operating convention; it is not an Adobe PDF-splitting limit. Reusing a small template for many exports is different from creating a single PBIX with every output page duplicated.

References: [Microsoft PDF export limitations](https://learn.microsoft.com/en-us/power-bi/collaborate-share/end-user-pdf) · [Adobe Acrobat OLE automation](https://opensource.adobe.com/dc-acrobat-sdk-docs/library/interapp/IAC_API_OLE_Objects.html).

## 中文说明

第一种：Power BI 先导出整份 PDF，再由 Excel VBA 调用 Adobe Acrobat 自动化库拆分、命名并保存。按本流程的使用要求，整份导出保持在 **50 页以下**。Adobe 原始 VBA 模块尚未找回，本次仅提供流程说明。

第二种：Excel VBA 调用 Windows UI Automation，模拟人工选择书签、调整页面可见性、导出并保存 PDF。可以理解为“仿真操作 agent”，但技术上属于固定规则的 RPA，**运行时消耗 0 模型 token**。需要保持 Power BI 打开、桌面解锁，运行时不要抢占鼠标、键盘或窗口焦点；界面更新、书签未保存正确、加载太慢等情况仍可能导致失败或内容错误。

源码、示例路径、部门/书签名称已改成通用 A/B/C/D。未上传实际报表、工作簿、内部路径、业务数据或原始日志。公开版尚未在 Windows 上重新编译和实跑；分发前需要检查日期、筛选条件和每页内容。

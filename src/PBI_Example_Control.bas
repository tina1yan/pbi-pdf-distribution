Attribute VB_Name = "PBI_Example_Control"
Option Explicit

Public Sub PBI_CreateExampleControl()
    Dim ws As Worksheet, existing As Worksheet
    Dim labels As Variant, headers As Variant, i As Long
    On Error Resume Next
    Set existing = ThisWorkbook.Worksheets("ABCD Control")
    On Error GoTo 0
    If Not existing Is Nothing Then
        MsgBox "ABCD Control already exists; no changes made.", vbInformation
        Exit Sub
    End If
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "ABCD Control"
    ws.Cells.NumberFormat = "@"
    ws.Range("A1").Value2 = "Synthetic example - configure a private test copy before running"
    ws.Range("A16").Value2 = "PBIX folder"
    ws.Range("B16").Value2 = "C:\abcd\input"
    ws.Range("A17").Value2 = "PBIX filename"
    ws.Range("B17").Value2 = "ABCD.pbix"
    headers = Array("Filename", "Folder", "Bookmark", "Volume Tab", "Run", "Status")
    For i = 0 To UBound(headers)
        ws.Cells(19, i + 1).Value2 = headers(i)
    Next i
    labels = Array("A", "B", "C", "D")
    For i = 0 To UBound(labels)
        ws.Cells(20 + i, 1).Value2 = labels(i) & "_yyyy-mm-dd.pdf"
        ws.Cells(20 + i, 2).Value2 = "C:\abcd\output"
        ws.Cells(20 + i, 3).Value2 = labels(i)
        If i < 3 Then ws.Cells(20 + i, 4).Value2 = "Volume " & labels(i)
        ws.Cells(20 + i, 5).Value2 = "No"
        ws.Cells(20 + i, 6).Value2 = "Not run"
    Next i
    ws.Rows(19).Font.Bold = True
    ws.Columns("A:F").AutoFit
    ws.Activate
    ws.Range("A20").Select
    MsgBox "Example created. Supply a test PBIX and folders, replace the filename date placeholder, then select one row and set Run=Yes. See docs/SETUP.md.", vbInformation
End Sub

Sub POGenerate()

Application.ScreenUpdating = False

Dim ats As Worksheet
Set ats = ThisWorkbook.Sheets("PO Macro")
Dim sourceName As Workbook
Dim vendorID As Integer

'clearing data
ats.Range("A9:AA5000").ClearContents

'Cell names
ats.Cells(8, 1) = "Posting date"
ats.Cells(8, 2) = "Material"
ats.Cells(8, 3) = "Material Name"
ats.Cells(8, 4) = "Manufacturer number"
ats.Cells(8, 5) = "PO number"
ats.Cells(8, 6) = "Intake Quantity"
ats.Cells(8, 7) = "Reference"
ats.Cells(8, 8) = "State"
ats.Cells(8, 9) = "Pallets"

ats.Cells(8, 10) = "State"
ats.Cells(9, 10) = "New"
ats.Cells(10, 10) = "Returned"
ats.Cells(11, 10) = "Total"

ats.Cells(8, 11) = "Total Quantity"
ats.Cells(8, 12) = "Pallets"

ats.Cells(8, 13) = "Number of POs"
ats.Cells(9, 13) = "=COUNTA(N:N)-1"

ats.Cells(8, 14) = "PO"

ats.Cells(5, 4) = "Error msg.:"
ats.Cells(6, 5) = "=COUNTA(F:F) - 1"
ats.Cells(4, 4) = "Vendor nr.:"
ats.Cells(6, 4) = "Line nr.:"

'Extra functions
'Color change
For x = 1 To 14

With ats.Cells(8, x).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
Next x

'Open direct file and find Posting Date
Workbooks.Open Filename:= _
        "K:\TEMPORARY\Andrius S\3PL\PO.xlsx"
ActiveSheet.Range("$A$1:$V$46456").AutoFilter Field:=17, Criteria1:= _
        ats.Cells(4, 5).Value
    Cells(2, 1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("MMFile.xlsm").Activate
    Sheets("PO Macro").Select
    Cells(9, 1).Select
    ats.Paste

'Material number
Workbooks("PO.xlsx").Activate
    Cells(2, 2).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("MMFile.xlsm").Activate
    Sheets("PO Macro").Select
    Cells(9, 2).Select
    ats.Paste

'Material Name
Workbooks("PO.xlsx").Activate
    Cells(2, 4).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("MMFile.xlsm").Activate
    Sheets("PO Macro").Select
    Cells(9, 3).Select
    ats.Paste

'PO
Workbooks("PO.xlsx").Activate
    Cells(2, 7).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("MMFile.xlsm").Activate
    Sheets("PO Macro").Select
    Cells(9, 5).Select
    ats.Paste

'QTY
Workbooks("PO.xlsx").Activate
        Cells(2, 20).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        Workbooks("MMFile.xlsm").Activate
        Sheets("PO Macro").Select
        Cells(9, 6).Select
        ats.Paste


'Reference
Workbooks("PO.xlsx").Activate
        Cells(2, 22).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        Workbooks("MMFile.xlsm").Activate
        Sheets("PO Macro").Select
        Cells(9, 7).Select
        ats.Paste

'Pasting PO and removing dublicates
Workbooks("PO.xlsx").Activate
    Cells(2, 7).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks("MMFile.xlsm").Activate
    Sheets("PO Macro").Select
    Cells(9, 14).Select
    ats.Paste
    ats.Range("$N$9:$N$2102").RemoveDuplicates Columns:=1, Header:= _
        xlYes
    
'Closing data file
Application.CutCopyMode = False
Workbooks("PO.xlsx").Close



End Sub



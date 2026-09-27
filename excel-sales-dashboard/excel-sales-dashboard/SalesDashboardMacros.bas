Attribute VB_Name = "SalesDashboardMacros"
'==========================================================================
' Sales Performance Dashboard - VBA macros
' Import: save the workbook as .xlsm, press Alt+F11, File > Import File,
' choose this .bas file, then assign the macros to buttons on the Dashboard.
'==========================================================================
Option Explicit

Private Const DASH_SHEET As String = "Dashboard"
Private Const DATA_SHEET As String = "Sales Data"
Private Const REGION_CELL As String = "C5"

' Refresh Power Query connections and recalculate every formula.
Public Sub RefreshDashboard()
    Application.ScreenUpdating = False
    On Error Resume Next
    ThisWorkbook.RefreshAll
    On Error GoTo 0
    Application.CalculateFull
    Application.ScreenUpdating = True
    MsgBox "Dashboard refreshed.", vbInformation
End Sub

' Show all regions again.
Public Sub ResetRegionFilter()
    ThisWorkbook.Worksheets(DASH_SHEET).Range(REGION_CELL).Value = "All"
End Sub

' Save the Dashboard as a PDF next to the workbook, e.g.
' Sales Dashboard - West - 2025-11-30.pdf
Public Sub ExportDashboardPDF()
    Dim ws As Worksheet, region As String, fileName As String
    Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    If ThisWorkbook.Path = "" Then
        MsgBox "Please save the workbook first, then run this macro again.", vbExclamation
        Exit Sub
    End If
    region = CStr(ws.Range(REGION_CELL).Value)
    fileName = ThisWorkbook.Path & Application.PathSeparator & _
               "Sales Dashboard - " & region & " - " & Format(Date, "yyyy-mm-dd") & ".pdf"
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fileName, _
        Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=True
End Sub

' Export one PDF per region in a single click.
Public Sub ExportAllRegionsPDF()
    Dim ws As Worksheet, regions As Variant, i As Long, original As String
    Set ws = ThisWorkbook.Worksheets(DASH_SHEET)
    If ThisWorkbook.Path = "" Then
        MsgBox "Please save the workbook first, then run this macro again.", vbExclamation
        Exit Sub
    End If
    original = CStr(ws.Range(REGION_CELL).Value)
    regions = Array("All", "North", "South", "East", "West")
    Application.ScreenUpdating = False
    For i = LBound(regions) To UBound(regions)
        ws.Range(REGION_CELL).Value = regions(i)
        Application.Calculate
        ws.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=ThisWorkbook.Path & Application.PathSeparator & _
                "Sales Dashboard - " & regions(i) & ".pdf", _
            Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
    Next i
    ws.Range(REGION_CELL).Value = original
    Application.ScreenUpdating = True
    MsgBox "Saved " & (UBound(regions) + 1) & " PDFs in " & ThisWorkbook.Path, vbInformation
End Sub

' Add a new order row with the next Order ID and today's date,
' copying the lookup formulas down from the row above.
Public Sub AddNewOrder()
    Dim ws As Worksheet, lastRow As Long, newRow As Long, lastId As String
    Set ws = ThisWorkbook.Worksheets(DATA_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    newRow = lastRow + 1
    lastId = CStr(ws.Cells(lastRow, "A").Value)

    ws.Cells(newRow, "A").Value = "SO-" & (CLng(Mid(lastId, 4)) + 1)
    ws.Cells(newRow, "B").Value = Date
    ws.Cells(newRow, "B").NumberFormat = "dd-mmm-yyyy"
    ' Copy formula columns: Month, Category, Unit Price, Revenue, Unit Cost, Cost, Profit
    ws.Range("C" & lastRow).Copy ws.Range("C" & newRow)
    ws.Range("G" & lastRow & ":G" & lastRow).Copy ws.Range("G" & newRow)
    ws.Range("I" & lastRow & ":M" & lastRow).Copy ws.Range("I" & newRow)

    ws.Activate
    ws.Cells(newRow, "D").Select
    MsgBox "New order row added (row " & newRow & ")." & vbCrLf & _
           "Fill in Region, Salesperson, Product and Units." & vbCrLf & vbCrLf & _
           "Note: dashboard formulas cover rows 2 to 571. Extend those ranges " & _
           "(or convert the data to an Excel Table) when you add many orders.", vbInformation
End Sub

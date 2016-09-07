Attribute VB_Name = "HelloWorld"
Option Explicit

Sub ClearAllFilters()
    On Error Resume Next
    ActiveSheet.ShowAllData
End Sub

Sub GraphicsToggleClear()
    ActiveSheet.Range("N3", Range("N3").End(xlDown)).AutoFilter Field:=13 'SO Personalized
End Sub

Sub GraphicsToggleYes()
    ActiveSheet.Range("N3", Range("N3").End(xlDown)).AutoFilter Field:=13, Criteria1:="Y" 'SO Personalized
End Sub

Sub GraphicsToggleNo()
    ActiveSheet.Range("N3", Range("N3").End(xlDown)).AutoFilter Field:=13, Criteria1:="N" 'SO Personalized
End Sub

Sub DTCOrderType()
    ClearAllFilters
    ActiveSheet.Range("Row3").AutoFilter Field:=29, Criteria1:="DTC Sales Order" 'Order Type
    ActiveSheet.Range("Row3").AutoFilter Field:=63, Criteria1:="1" 'Order Quantity
End Sub

Sub DTCOrderTypeClear()
    ActiveSheet.Range("Row3").AutoFilter Field:=29 'Order Type
    ActiveSheet.Range("Row3").AutoFilter Field:=63 'Order Quantity
End Sub

Sub DTCOrderTypeDuplicateRemoval()
    Dim DTCSheet As Object
    Range("K3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Select
    ActiveSheet.Name = "DTC Sales Orders"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Range("A:A").RemoveDuplicates Columns:=1, Header:=xlYes
End Sub

Sub AutoEligible100()
    ActiveSheet.Range("Row3").AutoFilter Field:=12, Criteria1:="100" 'Auto Eligible %
End Sub

Sub AutoEligibleNot100()
    ActiveSheet.Range("Row3").AutoFilter Field:=12, Criteria1:="<>100" 'Auto Eligible %
End Sub

Sub ComplianceLevel1()
    ActiveSheet.Range("Row3").AutoFilter Field:=26, Criteria1:="CC-1 (RG & EDI)" 'Compliance Level
End Sub

Sub ComplianceLevel2()
    ActiveSheet.Range("Row3").AutoFilter Field:=26, Criteria1:="CC-2 (RG)" 'Compliance Level
End Sub

Sub ComplianceLevel3()
    ActiveSheet.Range("Row3").AutoFilter Field:=26, Criteria1:="CC-3 (Non-Standard)" 'Compliance Level
End Sub

Sub ComplianceLevel4()
    ActiveSheet.Range("Row3").AutoFilter Field:=26, Criteria1:="CC-4 (Standard)" 'Compliance Level
End Sub

Sub ComplianceLevelBlank()
    ActiveSheet.Range("Row3").AutoFilter Field:=26, Criteria1:="=" 'Compliance Level
End Sub

Sub ComplianceLevelClear()
    ActiveSheet.Range("Row3").AutoFilter Field:=26 'Compliance Level
End Sub

Sub BatchNumberBlanks()
    ActiveSheet.Range("Row3").AutoFilter Field:=7, Criteria1:="="  'Batch #
End Sub

Sub BatchNumberClear()
    ActiveSheet.Range("Row3").AutoFilter Field:=7  'Batch #
End Sub

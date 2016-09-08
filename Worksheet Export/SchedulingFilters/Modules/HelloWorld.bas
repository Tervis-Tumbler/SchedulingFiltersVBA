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

Sub InvokeShipPriorityFilterSameDay()
    ActiveSheet.Range("Row3").AutoFilter Field:=55, Criteria1:="Same Day Rush" 'Ship Priority
End Sub
Sub InvokeShipPriorityFilterRushNotSameDay()
    ActiveSheet.Range("Row3").AutoFilter Field:=55, Criteria1:=Array( _
            "Rush 1D", "Rush 2D", "Rush 3D" _
        ) 'Ship Priority
End Sub
Sub InvokeShipPriorityFilterStandard()
    ActiveSheet.Range("Row3").AutoFilter Field:=55, Criteria1:="Standard" 'Ship Priority
End Sub
Sub InvokeShipPriorityClearFilter()
    ActiveSheet.Range("Row3").AutoFilter Field:=55 'Ship Priority
End Sub


Sub TestNewEvent()
    NewEvent Message:="{FunctionName:'Value'}"
End Sub

Sub TestNewSchedlingFilterFunctionCallEvent()
    NewSchedlingFilterFunctionCallEvent FunctionName:="TestFunctionName"

End Sub

Sub DTCDuplicateRemovalSheetPersonalized()
    Application.ScreenUpdating = False
    Dim DTCSheet As Object
    Range("K3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Select
    ActiveSheet.Name = "Personalized"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Range("A:A").RemoveDuplicates Columns:=1, Header:=xlYes
    Application.ScreenUpdating = True
End Sub

Sub DTCDuplicateRemovalSheetPersonalized1Cup()
    Application.ScreenUpdating = False
    Dim DTCSheet As Object
    Range("K3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Select
    ActiveSheet.Name = "Personalized, 1 Cup"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Range("A:A").RemoveDuplicates Columns:=1, Header:=xlYes
    Application.ScreenUpdating = True
End Sub

Sub DTCDuplicateRemovalSheetAutoEligible()
    Application.ScreenUpdating = False
    Dim DTCSheet As Object
    Range("K3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Select
    ActiveSheet.Name = "Auto Eligible"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Range("A:A").RemoveDuplicates Columns:=1, Header:=xlYes
    Application.ScreenUpdating = True
End Sub

Sub DTCDuplicateRemovalSheetNotPersonalized()
    Application.ScreenUpdating = False
    Dim DTCSheet As Object
    Range("K3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Select
    ActiveSheet.Name = "Not Personalized"
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Range("A:A").RemoveDuplicates Columns:=1, Header:=xlYes
    Application.ScreenUpdating = True
End Sub

Sub DTCAll()
    DTCFilter4
    DTCDuplicateRemovalSheet4
    DTCFilter3
    DTCDuplicateRemovalSheet3
    DTCFilter2
    DTCDuplicateRemovalSheet2
    DTCFilter1
    DTCDuplicateRemovalSheet1
End Sub

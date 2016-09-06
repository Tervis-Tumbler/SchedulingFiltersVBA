Attribute VB_Name = "DTC"
Enum OrderTypeEnum
    DTC
    Warranty
End Enum

Sub InvokeSchedulingFilter _
    ( _
        OrderType As OrderTypeEnum, _
        OneHundredPercentAutoEligible As Boolean, _
        SOPersonalized As Boolean, _
        Optional FilterForOrderQuantityEqualOne As Boolean = False, _
        Optional ShipPrioritySameDayRush As Boolean = False, _
        Optional ShipPriorityRushNotSameDay As Boolean = False, _
        Optional ShipPriorityStandard As Boolean = False _
    )
    
    Application.ScreenUpdating = False
    Worksheets("Sheet1").Activate
    ClearAllFilters
    
    Select Case OrderType
        Case OrderTypeEnum.DTC
            ActiveSheet.Range("Row3").AutoFilter Field:=29, Criteria1:="DTC Sales Order" 'Order Type
        Case OrderTypeEnum.Warranty
            ActiveSheet.Range("Row3").AutoFilter Field:=29, Criteria1:="Warranty Order" 'Order Type
    End Select
    
    If SOPersonalized Then
        ActiveSheet.Range("N3", Range("N3").End(xlDown)).AutoFilter Field:=13, Criteria1:="Y" 'SO Personalized
    Else
        ActiveSheet.Range("N3", Range("N3").End(xlDown)).AutoFilter Field:=13, Criteria1:="N" 'SO Personalized
    End If
    
    If OneHundredPercentAutoEligible Then
        ActiveSheet.Range("Row3").AutoFilter Field:=12, Criteria1:="100" 'Auto Eligible %
    Else
        ActiveSheet.Range("Row3").AutoFilter Field:=12, Criteria1:="<>100" 'Auto Eligible %
    End If
    
    If ShipPrioritySameDayRush Then
        InvokeShipPriorityFilterSameDay
    End If
    
    If ShipPriorityRushNotSameDay Then
        InvokeShipPriorityFilterRushNotSameDay
    End If
    
    If ShipPriorityStandard Then
        InvokeShipPriorityFilterStandard
    End If
    
    If OrderQuantity Then
        ActiveSheet.Range("Row3").AutoFilter Field:=63, Criteria1:="1" 'Order Quantity
    End If
    
    Application.ScreenUpdating = True
End Sub

Sub InvokeDTCFilterPersonalized()
    InvokeSchedulingFilter OrderType:=OrderTypeEnum.DTC, SOPersonalized:=True, OneHundredPercentAutoEligible:=False, ShipPriorityStandard:=True
    NewSchedlingFilterFunctionCallEvent FunctionName:="InvokeDTCFilterPersonalized"
End Sub
Sub InvokeDTCFilterNotPersonalized1Cup()
    InvokeSchedulingFilter OrderType:=OrderTypeEnum.DTC, SOPersonalized:=False, OneHundredPercentAutoEligible:=False, FilterForOrderQuantityEqualOne:=True, ShipPriorityStandard:=True
    NewSchedlingFilterFunctionCallEvent FunctionName:="InvokeDTCFilterNotPersonalized1Cup"
End Sub
Sub InvokeDTCFilterAutoEligible()
    InvokeSchedulingFilter OrderType:=OrderTypeEnum.DTC, SOPersonalized:=False, OneHundredPercentAutoEligible:=True, ShipPriorityStandard:=True
    NewSchedlingFilterFunctionCallEvent FunctionName:="InvokeDTCFilterAutoEligible"
End Sub
Sub InvokeDTCFilterNotPersonalized()
    InvokeSchedulingFilter OrderType:=OrderTypeEnum.DTC, SOPersonalized:=False, OneHundredPercentAutoEligible:=False, ShipPriorityStandard:=True
    NewSchedlingFilterFunctionCallEvent FunctionName:="InvokeDTCFilterNotPersonalized"
End Sub

Sub InvokeWarrantyFilterPersonalized()
    InvokeSchedulingFilter OrderType:=OrderTypeEnum.Warranty, SOPersonalized:=True, OneHundredPercentAutoEligible:=False
    NewSchedlingFilterFunctionCallEvent FunctionName:="InvokeWarrantyFilterPersonalized"
End Sub
Sub InvokeWarrantyFilterNotPersonalized1Cup()
    InvokeSchedulingFilter OrderType:=OrderTypeEnum.Warranty, SOPersonalized:=False, OneHundredPercentAutoEligible:=False, FilterForOrderQuantityEqualOne:=True
    NewSchedlingFilterFunctionCallEvent FunctionName:="InvokeWarrantyFilterNotPersonalized1Cup"
End Sub
Sub InvokeWarrantyFilterAutoEligible()
    InvokeSchedulingFilter OrderType:=OrderTypeEnum.Warranty, SOPersonalized:=False, OneHundredPercentAutoEligible:=True
    NewSchedlingFilterFunctionCallEvent FunctionName:="InvokeWarrantyFilterAutoEligible"
End Sub
Sub InvokeWarrantyFilterNotPersonalized()
    InvokeSchedulingFilter OrderType:=OrderTypeEnum.Warranty, SOPersonalized:=False, OneHundredPercentAutoEligible:=False
    NewSchedlingFilterFunctionCallEvent FunctionName:="InvokeWarrantyFilterNotPersonalized"
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
Sub InvokeDuplicateRemoval()

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

Sub NewEvent(Message As String)
    Set WSScriptShell = CreateObject("WScript.Shell")
    WSScriptShell.LogEvent 4, Message
End Sub

Sub NewSchedlingFilterFunctionCallEvent(FunctionName As String)
    NewEvent Message:="{FunctionName:'" & FunctionName & "'}"
End Sub

Sub TestNewEvent()
    NewEvent Message:="{FunctionName:'Value'}"
End Sub

Sub TestNewSchedlingFilterFunctionCallEvent()
    NewSchedlingFilterFunctionCallEvent FunctionName:="TestFunctionName"

End Sub

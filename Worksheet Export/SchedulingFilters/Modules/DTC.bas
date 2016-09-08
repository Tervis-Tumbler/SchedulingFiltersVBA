Attribute VB_Name = "DTC"
Enum OrderTypeEnum
    DTC
    Warranty
End Enum

Enum OrderSubtypeEnum
    DTC_Personalized
    DTC_Not_Personalized
    DTC_Not_Personlized_1_Cup
    DTC_Auto_Eligible
    Warranty_Personalized
    Warranty_Not_Personalized
    Warranty_Not_Personlized_1_Cup
    Warranty_Auto_Eligible
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

Sub InvokeDuplicateRemoval( _
    OrderSubtype As OrderSubtypeEnum _
    )
    
    Application.ScreenUpdating = False
    Range("K3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Select
    
    Select Case OrderSubtype
        Case OrderSubtypeEnum.DTC_Auto_Eligible
            ActiveSheet.Name = "DTC Auto Eligible"
        Case OrderSubtypeEnum.DTC_Not_Personalized
            AtiveSheet.Name = "DTC Not Personalized"
        Case OrderSubtypeEnum.DTC_Not_Personlized_1_Cup
            ActiveSheet.Name = "DTC Not Personalized 1Cup"
        Case OrderSubtypeEnum.DTC_Personalized
            ActiveSheet.Name = "DTC_Personalized"
        Case OrderSubtypeEnum.Warranty_Auto_Eligible
            ActiveSheet.Name = "Warranty Auto Eligible"
        Case OrderSubtypeEnum.Warranty_Not_Personalized
            ActiveSheet.Name = "Warranty Not Personalized"
        Case OrderSubtypeEnum.Warranty_Not_Personlized_1_Cup
            ActiveSheet.Name = "Warranty Not Personalized 1Cup"
        Case OrderSubtypeEnum.Warranty_Personalized
            ActiveSheet.Name = "Warranty Personalized"
    End Select
            
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Range("A:A").RemoveDuplicates Columns:=1, Header:=xlYes
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

Sub InvokeDTCDuplicateRemovalPersonalized()
    InvokeDuplicateRemoval OrderSubtype:=OrderSubtypeEnum.DTC_Personalized
    NewSchedlingFilterFunctionCallEvent FunctionName:="InvokeDTCDuplicateRemovalPersonalized"
End Sub
Sub InvokeDTCDuplicateRemovalNotPersonalized()
    InvokeDuplicateRemoval OrderSubtype:=OrderSubtypeEnum.DTC_Not_Personalized
    NewSchedlingFilterFunctionCallEvent FunctionName:="InvokeDTCDuplicateRemovalNotPersonalized"
End Sub
Sub InvokeDTCDuplicateRemovalNotPersonalized1Cup()
    InvokeDuplicateRemoval OrderSubtype:=OrderSubtypeEnum.DTC_Not_Personlized_1_Cup
    NewSchedlingFilterFunctionCallEvent FunctionName:="InvokeDTCDuplicateRemovalNotPersonalized1Cup"
End Sub
Sub InvokeDTCDuplicateRemovalAutoEligible()
    InvokeDuplicateRemoval OrderSubtype:=OrderSubtypeEnum.DTC_Auto_Eligible
    NewSchedlingFilterFunctionCallEvent FunctionName:="InvokeDTCDuplicateRemovalAutoEligible"
End Sub
Sub InvokeWarrantyDuplicateRemovalPersonalized()
    InvokeDuplicateRemoval OrderSubtype:=OrderSubtypeEnum.Warranty_Personalized
    NewSchedlingFilterFunctionCallEvent FunctionName:="InvokeWarrantyDuplicateRemovalPersonalized"
End Sub
Sub InvokeWarrantyDuplicateRemovalNotPersonalized()
    InvokeDuplicateRemoval OrderSubtype:=OrderSubtypeEnum.Warranty_Not_Personalized
    NewSchedlingFilterFunctionCallEvent FunctionName:="InvokeWarrantyDuplicateRemovalNotPersonalized"
End Sub
Sub InvokeWarrantyDuplicateRemovalNotPersonalized1Cup()
    InvokeDuplicateRemoval OrderSubtype:=OrderSubtypeEnum.Warranty_Not_Personlized_1_Cup
    NewSchedlingFilterFunctionCallEvent FunctionName:="InvokeWarrantyDuplicateRemovalNotPersonalized1Cup"
End Sub
Sub InvokeWarrantyDuplicateRemovalAutoEligible()
    InvokeDuplicateRemoval OrderSubtype:=OrderSubtypeEnum.Warranty_Auto_Eligible
    NewSchedlingFilterFunctionCallEvent FunctionName:="InvokeWarrantyDuplicateRemovalAutoEligible"
End Sub

Sub NewEvent(Message As String)
    Set WSScriptShell = CreateObject("WScript.Shell")
    WSScriptShell.LogEvent 4, Message
End Sub

Sub NewSchedlingFilterFunctionCallEvent(FunctionName As String)
    NewEvent Message:="{FunctionName:'" & FunctionName & "'}"
End Sub

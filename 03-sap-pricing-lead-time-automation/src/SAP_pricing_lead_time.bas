
Option Explicit

' ============================================================
' SAP Pricing & Lead Time Automation
'
' Portfolio Version
'
' Workflow:
' Excel Input
'      ↓
' SAP VA01
'      ↓
' Pricing & Lead Time Lookup
'      ↓
' Excel Update
'
' This public portfolio version preserves the core workflow
' while removing company-specific workbook layouts and values.
' ============================================================

Private Const SHEET_NAME As String = "Demo"
Private Const START_ROW As Long = 5

' Demo workbook columns
Private Const COL_MATERIAL As Long = 1       ' A
Private Const COL_QTY As Long = 2            ' B
Private Const COL_LIST As Long = 5           ' E
Private Const COL_NET As Long = 6            ' F
Private Const COL_DISCOUNT As Long = 7       ' G
Private Const COL_PLANT As Long = 8          ' H
Private Const COL_EST_DATE As Long = 9       ' I
Private Const COL_WEEK As Long = 10          ' J

' SAP GUI IDs used by the automation
Private Const ID_OKCD As String = "wnd[0]/tbar[0]/okcd"
Private Const ID_AUART As String = "wnd[0]/usr/ctxtVBAK-AUART"
Private Const ID_VKORG As String = "wnd[0]/usr/ctxtVBAK-VKORG"
Private Const ID_VTWEG As String = "wnd[0]/usr/ctxtVBAK-VTWEG"

Private Const ID_OVERVIEW_TAB As String = _
    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01"

Private Const ID_ITEM_COND_TAB As String = _
    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06"

Private Const ID_TBL_OVERVIEW As String = _
    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/" & _
    "ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/" & _
    "tblSAPMV45ATCTRL_U_ERF_AUFTRAG"

Private Const ID_TBL_COND As String = _
    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/" & _
    "ssubSUBSCREEN_BODY:SAPLV69A:6201/" & _
    "tblSAPLV69ATCTRL_KONDITIONEN"

Public Sub RunPricingLeadTimeLookup()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim itemIndex As Long
    Dim material As String
    Dim qty As String

    Dim SapGuiAuto As Object
    Dim SapApp As Object
    Dim connection As Object
    Dim session As Object

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    lastRow = ws.Cells(ws.Rows.Count, COL_MATERIAL).End(xlUp).Row
    If lastRow < START_ROW Then Exit Sub

    ' Attach to an existing SAP GUI session.
    On Error GoTo SAP_CONNECTION_ERROR

    Set SapGuiAuto = GetObject("SAPGUI")
    Set SapApp = SapGuiAuto.GetScriptingEngine
    Set connection = SapApp.Children(0)
    Set session = connection.Children(0)

    session.findById("wnd[0]").maximize

    ' Open VA01.
    session.findById(ID_OKCD).Text = "/nva01"
    session.findById("wnd[0]").sendVKey 0
    CloseSapPopups session

    ' Portfolio placeholders.
    ' Replace these values with the appropriate values for your SAP environment.
    session.findById(ID_AUART).Text = "<ORDER_TYPE>"
    session.findById(ID_VKORG).Text = "<SALES_ORG>"
    session.findById(ID_VTWEG).Text = "<DISTRIBUTION_CHANNEL>"
    session.findById("wnd[0]").sendVKey 0
    CloseSapPopups session

    ' --------------------------------------------------------
    ' Enter Material and Quantity from Excel
    ' --------------------------------------------------------
    For r = START_ROW To lastRow

        material = Trim$(CStr(ws.Cells(r, COL_MATERIAL).Value))
        qty = Trim$(CStr(ws.Cells(r, COL_QTY).Value))

        If material = "" Then Exit For

        itemIndex = r - START_ROW

        session.findById(ID_OVERVIEW_TAB).Select

        session.findById( _
            ID_TBL_OVERVIEW & "/ctxtRV45A-MABNR[1," & itemIndex & "]").Text = material

        session.findById( _
            ID_TBL_OVERVIEW & "/txtRV45A-KWMENG[2," & itemIndex & "]").Text = qty

        session.findById("wnd[0]").sendVKey 0
        CloseSapPopups session

    Next r

    ' --------------------------------------------------------
    ' Retrieve Pricing and Lead Time Information
    ' --------------------------------------------------------
    For r = START_ROW To lastRow

        material = Trim$(CStr(ws.Cells(r, COL_MATERIAL).Value))
        If material = "" Then Exit For

        itemIndex = r - START_ROW

        ' Open item details.
        session.findById( _
            ID_TBL_OVERVIEW & "/ctxtRV45A-MABNR[1," & itemIndex & "]").SetFocus

        session.findById("wnd[0]").sendVKey 2
        CloseSapPopups session

        ' Pricing conditions.
        session.findById(ID_ITEM_COND_TAB).Select
        CloseSapPopups session

        ws.Cells(r, COL_LIST).Value = GetConditionValue(session, "ZA01")
        ws.Cells(r, COL_NET).Value = GetConditionValue(session, "NETP")

        Dim discountRaw As String
        discountRaw = GetConditionValue(session, "YA9I")

        If discountRaw = "" Then
            discountRaw = GetConditionValue(session, "Y18R")
        End If

        If discountRaw = "" Then discountRaw = "0"

        ws.Cells(r, COL_DISCOUNT).Value = NormalizeDiscount(discountRaw) / 100
        ws.Cells(r, COL_DISCOUNT).NumberFormat = "0.00%"

        ' Return to overview.
        session.findById("wnd[0]/tbar[0]/btn[3]").press
        CloseSapPopups session

        ' Plant.
        ws.Cells(r, COL_PLANT).Value = _
            GetOverviewValue(session, "VBAP-WERKS", itemIndex)

        ' Estimated delivery date.
        Dim deliveryDateText As String
        Dim deliveryDate As Date

        deliveryDateText = _
            GetOverviewValue(session, "VBAP-ZZRSTD", itemIndex)

        ws.Cells(r, COL_EST_DATE).Value = deliveryDateText

        ' Original workflow creates the delivery-week label
        ' from the SAP estimated delivery date plus 7 days.
        If TryParseDate(deliveryDateText, deliveryDate) Then
            ws.Cells(r, COL_WEEK).Value = _
                WeekLabel(DateAdd("d", 7, deliveryDate))
        Else
            ws.Cells(r, COL_WEEK).Value = ""
        End If

    Next r

    MsgBox "Pricing and lead time lookup completed.", vbInformation
    Exit Sub

SAP_CONNECTION_ERROR:
    MsgBox "Unable to connect to SAP." & vbCrLf & _
           "Please make sure SAP GUI is open and logged in.", vbCritical

End Sub

Private Function GetConditionValue( _
    ByVal session As Object, _
    ByVal conditionType As String) As String

    Dim i As Long
    Dim currentType As String
    Dim result As String

    On Error Resume Next

    For i = 0 To 40

        currentType = session.findById( _
            ID_TBL_COND & "/ctxtKOMV-KSCHL[0," & i & "]").Text

        If UCase$(Trim$(currentType)) = _
           UCase$(Trim$(conditionType)) Then

            result = session.findById( _
                ID_TBL_COND & "/txtKOMV-KWERT[3," & i & "]").Text

            If Trim$(result) = "" Then
                result = session.findById( _
                    ID_TBL_COND & "/txtKOMV-KBETR[3," & i & "]").Text
            End If

            GetConditionValue = Trim$(result)
            Exit Function

        End If

    Next i

    On Error GoTo 0
    GetConditionValue = ""

End Function

Private Function GetOverviewValue( _
    ByVal session As Object, _
    ByVal fieldName As String, _
    ByVal itemIndex As Long) As String

    Dim col As Long
    Dim value As String
    Dim prefix As String

    If fieldName = "VBAP-WERKS" Or fieldName = "VBAP-ZZRSTD" Then
        prefix = "ctxt"
    Else
        prefix = "txt"
    End If

    On Error Resume Next

    ' Column positions may vary slightly by SAP layout.
    For col = 6 To 10

        value = session.findById( _
            ID_TBL_OVERVIEW & "/" & prefix & fieldName & _
            "[" & col & "," & itemIndex & "]").Text

        If Trim$(value) <> "" Then
            GetOverviewValue = Trim$(value)
            Exit Function
        End If

    Next col

    On Error GoTo 0
    GetOverviewValue = ""

End Function

Private Function NormalizeDiscount(ByVal value As String) As Double

    Dim textValue As String

    textValue = Trim$(value)
    textValue = Replace(textValue, " ", "")
    textValue = Replace(textValue, ",", "")
    textValue = Replace(textValue, "%", "")

    If Len(textValue) > 0 And Right$(textValue, 1) = "-" Then
        textValue = Left$(textValue, Len(textValue) - 1)
    End If

    If textValue = "" Then
        NormalizeDiscount = 0
        Exit Function
    End If

    On Error GoTo INVALID_VALUE

    NormalizeDiscount = Abs(CDbl(textValue))
    Exit Function

INVALID_VALUE:
    NormalizeDiscount = 0

End Function

Private Function TryParseDate( _
    ByVal value As String, _
    ByRef outputDate As Date) As Boolean

    Dim parts() As String
    Dim textValue As String

    textValue = Trim$(value)

    If textValue = "" Then Exit Function

    On Error GoTo INVALID_DATE

    If InStr(textValue, ".") > 0 Then

        parts = Split(textValue, ".")
        outputDate = DateSerial( _
            CInt(parts(2)), CInt(parts(1)), CInt(parts(0)))

    ElseIf InStr(textValue, "/") > 0 Then

        parts = Split(textValue, "/")
        outputDate = DateSerial( _
            CInt(parts(2)), CInt(parts(1)), CInt(parts(0)))

    ElseIf InStr(textValue, "-") > 0 Then

        parts = Split(textValue, "-")

        If Len(parts(0)) = 4 Then
            outputDate = DateSerial( _
                CInt(parts(0)), CInt(parts(1)), CInt(parts(2)))
        Else
            outputDate = DateSerial( _
                CInt(parts(2)), CInt(parts(1)), CInt(parts(0)))
        End If

    Else
        outputDate = CDate(textValue)
    End If

    TryParseDate = True
    Exit Function

INVALID_DATE:
    TryParseDate = False

End Function

Private Function WeekLabel(ByVal value As Date) As String

    Dim weekNumber As Long

    weekNumber = Int((Day(value) - 1) / 7) + 1

    WeekLabel = _
        MonthName(Month(value)) & " Week " & weekNumber

End Function

Private Sub CloseSapPopups( _
    ByVal session As Object, _
    Optional ByVal maxLoops As Long = 10)

    Dim i As Long

    On Error Resume Next

    For i = 1 To maxLoops

        If session.Children.Count <= 1 Then Exit For

        session.findById("wnd[1]/tbar[0]/btn[0]").press

        If session.Children.Count <= 1 Then Exit For

        session.findById("wnd[1]").sendVKey 0

        If session.Children.Count <= 1 Then Exit For

        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

    Next i

    On Error GoTo 0

End Sub

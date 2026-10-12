
Option Explicit

' ============================================================
' SAP Invoice Data Extraction Automation
'
' Portfolio Version
'
' Workflow:
' Excel Delivery Number
'       ↓
' SAP VL03N
'       ↓
' Document Flow
'       ↓
' Related Invoice
'       ↓
' Invoice Number / Date / Net Value
'       ↓
' Excel Update
'
' Demo workbook mapping:
' A = Delivery Number
' C = Invoice Number
' D = Invoice Date
' E = Net Value
' F = Status
'
' Data starts from Row 5.
' ============================================================

Private Const SHEET_NAME As String = "Demo"
Private Const START_ROW As Long = 5

Private Const COL_DELIVERY As Long = 1
Private Const COL_INVOICE As Long = 3
Private Const COL_INV_DATE As Long = 4
Private Const COL_NET_VALUE As Long = 5
Private Const COL_STATUS As Long = 6

Public Sub RunInvoiceDataExtraction()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Dim delivery As String
    Dim invoiceNo As String
    Dim netValue As String
    Dim invoiceDate As String
    Dim parsedDate As Date
    Dim success As Boolean

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    Application.ScreenUpdating = False
    Application.StatusBar = "Retrieving invoice data from SAP..."

    lastRow = ws.Cells(ws.Rows.Count, COL_DELIVERY).End(xlUp).Row

    For r = START_ROW To lastRow

        delivery = Trim$(CStr(ws.Cells(r, COL_DELIVERY).Value))

        If delivery <> "" _
           And Trim$(CStr(ws.Cells(r, COL_INVOICE).Value)) = "" _
           And IsValidDelivery(delivery) Then

            invoiceNo = ""
            netValue = ""
            invoiceDate = ""

            Application.StatusBar = _
                "Processing Delivery: " & delivery

            success = GetInvoiceFromDelivery( _
                delivery, invoiceNo, netValue, invoiceDate)

            If success Then

                ws.Cells(r, COL_INVOICE).Value = invoiceNo
                ws.Cells(r, COL_NET_VALUE).Value = netValue

                If invoiceDate <> "" Then

                    parsedDate = ParseSapDate(invoiceDate)

                    If parsedDate > 0 Then
                        ws.Cells(r, COL_INV_DATE).Value = parsedDate
                        ws.Cells(r, COL_INV_DATE).NumberFormat = "dd-mmm-yy"
                    Else
                        ws.Cells(r, COL_INV_DATE).Value = invoiceDate
                    End If

                End If

                ws.Cells(r, COL_STATUS).Value = "Updated"

            Else

                ws.Cells(r, COL_STATUS).Value = "Invoice Not Found"

            End If

            DoEvents

        ElseIf delivery <> "" _
            And Trim$(CStr(ws.Cells(r, COL_INVOICE).Value)) <> "" Then

            ws.Cells(r, COL_STATUS).Value = "Existing Invoice - Skipped"

        ElseIf delivery <> "" And Not IsValidDelivery(delivery) Then

            ws.Cells(r, COL_STATUS).Value = "Invalid Delivery"

        End If

    Next r

    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox "Invoice data extraction completed.", vbInformation

End Sub

' ============================================================
' Validate Delivery Number
' Original workflow: 9-digit numeric value beginning with 4.
' ============================================================

Private Function IsValidDelivery(ByVal value As String) As Boolean

    value = Trim$(value)

    If Not IsNumeric(value) Then Exit Function

    IsValidDelivery = _
        (Left$(value, 1) = "4" And Len(value) = 9)

End Function

' ============================================================
' SAP Delivery → Related Invoice
'
' VL03N → Document Flow → Find Invoice → Display Invoice
' → Read Invoice Number / Net Value / Billing Date
' ============================================================

Private Function GetInvoiceFromDelivery( _
    ByVal delivery As String, _
    ByRef invoiceNo As String, _
    ByRef netValue As String, _
    ByRef invoiceDate As String) As Boolean

    Dim session As Object
    Dim tree As Object
    Dim invoiceNode As String

    Dim root As Object
    Dim invoiceField As Object
    Dim netValueField As Object
    Dim dateField As Object

    Dim startTime As Single

    On Error GoTo ERROR_HANDLER

    Set session = GetSapSession()

    If session Is Nothing Then Exit Function

    ' Open VL03N.
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nvl03n"
    session.findById("wnd[0]/tbar[0]/btn[0]").Press

    ' Enter Delivery Number.
    session.findById("wnd[0]/usr/ctxtLIKP-VBELN").Text = delivery
    session.findById("wnd[0]/usr/ctxtLIKP-VBELN").SetFocus
    session.findById("wnd[0]/tbar[1]/btn[8]").Press

    ' Environment → Document Flow.
    session.findById("wnd[0]/mbar/menu[4]/menu[7]").Select

    ' Find the related Invoice node.
    session.findById( _
        "wnd[0]/usr/shell/shellcont[1]/shell[0]").PressButton "&FIND"

    session.findById("wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "invoice"
    session.findById("wnd[1]/tbar[0]/btn[0]").Press
    session.findById("wnd[1]").Close

    Set tree = session.findById( _
        "wnd[0]/usr/shell/shellcont[1]/shell[1]")

    On Error Resume Next
    invoiceNode = tree.SelectedNode
    On Error GoTo ERROR_HANDLER

    If invoiceNode = "" Then Exit Function

    tree.SelectItem invoiceNode, "&Hierarchy"
    tree.EnsureVisibleHorizontalItem invoiceNode, "&Hierarchy"
    tree.SetFocus

    startTime = Timer

    Do While Timer - startTime < 0.3
        DoEvents
    Loop

    ' Display the selected Invoice document.
    On Error Resume Next

    session.findById("wnd[0]/tbar[1]/btn[8]").Press

    If Err.Number <> 0 Then
        Err.Clear
        session.findById("wnd[0]").sendVKey 8
    End If

    On Error GoTo ERROR_HANDLER

    Set root = session.findById("wnd[0]/usr")

    ' Wait briefly for the Invoice screen.
    Set invoiceField = Nothing
    startTime = Timer

    Do

        On Error Resume Next

        Set invoiceField = _
            root.findByName("VBRK-VBELN", "GuiCTextField")

        If invoiceField Is Nothing Then
            Set invoiceField = _
                root.findByName("VBRK-VBELN", "GuiTextField")
        End If

        On Error GoTo ERROR_HANDLER

        If Not invoiceField Is Nothing Then Exit Do

        DoEvents

    Loop While Timer - startTime < 5

    If invoiceField Is Nothing Then Exit Function

    ' Retrieve Net Value and Billing Date.
    On Error Resume Next

    Set netValueField = _
        root.findByName("VBRK-NETWR", "GuiTextField")

    If netValueField Is Nothing Then
        Set netValueField = _
            root.findByName("VBRK-NETWR", "GuiCTextField")
    End If

    Set dateField = _
        root.findByName("VBRK-FKDAT", "GuiTextField")

    If dateField Is Nothing Then
        Set dateField = _
            root.findByName("VBRK-FKDAT", "GuiCTextField")
    End If

    On Error GoTo ERROR_HANDLER

    invoiceNo = Trim$(invoiceField.Text)

    If Not netValueField Is Nothing Then
        netValue = Trim$(netValueField.Text)
    End If

    If Not dateField Is Nothing Then
        invoiceDate = Trim$(dateField.Text)
    End If

    GetInvoiceFromDelivery = True
    Exit Function

ERROR_HANDLER:
    GetInvoiceFromDelivery = False

End Function

' ============================================================
' Attach to an existing SAP GUI session.
' ============================================================

Private Function GetSapSession() As Object

    Dim SapGuiAuto As Object
    Dim SapApp As Object
    Dim connection As Object
    Dim session As Object

    On Error GoTo ERROR_HANDLER

    Set SapGuiAuto = GetObject("SAPGUI")
    Set SapApp = SapGuiAuto.GetScriptingEngine
    Set connection = SapApp.Children(0)
    Set session = connection.Children(0)

    Set GetSapSession = session
    Exit Function

ERROR_HANDLER:
    Set GetSapSession = Nothing

End Function

' ============================================================
' Convert SAP dd.mm.yyyy date text to VBA Date.
' ============================================================

Private Function ParseSapDate(ByVal value As String) As Date

    value = Trim$(value)

    On Error GoTo INVALID_DATE

    If Len(value) >= 10 Then

        ParseSapDate = DateSerial( _
            CInt(Mid$(value, 7, 4)), _
            CInt(Mid$(value, 4, 2)), _
            CInt(Left$(value, 2)))

    Else
        ParseSapDate = 0
    End If

    Exit Function

INVALID_DATE:
    ParseSapDate = 0

End Function

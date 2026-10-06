
Option Explicit

' =============================================================================
' SAP Delivery Date Automation - By Line Item
' SAP ERP/ECC | Excel VBA | SAP GUI Scripting
'
' PURPOSE
' -------
' Updates requested delivery dates in SAP ECC by locating the requested
' Line Item and validating Part Number and Quantity before making a change.
'
' EXCEL INPUT
' -----------
' B = Customer PO
' C = Line Item
' D = Part Number
' E = Quantity
' F = New Delivery Date
'
' OUTPUT
' ------
' H = SAP Part Number
' I = SAP Quantity
' K = Status
'
' SAFETY
' ------
' Part Number AND Quantity must match SAP.
' Match    -> Delivery date updated
' Mismatch -> No update
'
' First data row = 4
'
' Portfolio version - company/customer data removed.
' SAP GUI element IDs may vary between SAP environments.
' =============================================================================

Public Sub UpdateDeliveryDate_ByLineItem()

    Dim Sap As Object
    Dim App As Object
    Dim Conn As Object
    Dim Sess As Object

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Dim po As String
    Dim lineItem As String
    Dim expectedPN As String
    Dim expectedQty As Double
    Dim newDate As String

    Dim gridId As String

    Dim sapPN As String
    Dim sapQtyRaw As String
    Dim sapQty As Double

    Dim sapPNNorm As String
    Dim excelPNNorm As String

    Dim pnMatch As Boolean
    Dim qtyMatch As Boolean

    Dim cellId As String


    ' -------------------------------------------------------------------------
    ' SAP Item Overview table
    ' -------------------------------------------------------------------------

    gridId = _
        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/" & _
        "ssubSUBSCREEN_BODY:SAPMV45A:4400/" & _
        "subSUBSCREEN_TC:SAPMV45A:4900/" & _
        "tblSAPMV45ATCTRL_U_ERF_AUFTRAG"


    ' -------------------------------------------------------------------------
    ' Connect to SAP GUI
    ' -------------------------------------------------------------------------

    On Error GoTo SAPConnectionError

    Set Sap = GetObject("SAPGUI")
    Set App = Sap.GetScriptingEngine
    Set Conn = App.Children(0)
    Set Sess = Conn.Children(0)

    On Error GoTo 0


    ' -------------------------------------------------------------------------
    ' Excel worksheet
    ' -------------------------------------------------------------------------

    Set ws = ThisWorkbook.Sheets(1)

    lastRow = _
        ws.Cells(ws.Rows.Count, "B").End(xlUp).Row


    ' =========================================================================
    ' Process each request
    ' =========================================================================

    For r = 4 To lastRow

        po = Trim(CStr(ws.Cells(r, "B").Value))
        lineItem = Trim(CStr(ws.Cells(r, "C").Value))
        expectedPN = Trim(CStr(ws.Cells(r, "D").Value))


        ' Clear previous results

        ws.Cells(r, "H").Value = ""
        ws.Cells(r, "I").Value = ""
        ws.Cells(r, "K").Value = ""


        ' ---------------------------------------------------------------------
        ' Validate required input
        ' ---------------------------------------------------------------------

        If po = "" _
            Or lineItem = "" _
            Or expectedPN = "" _
            Or ws.Cells(r, "E").Value = "" Then

            ws.Cells(r, "K").Value = "Missing data"
            GoTo NextRow

        End If


        If Not IsNumeric(ws.Cells(r, "E").Value) Then

            ws.Cells(r, "K").Value = "Invalid Qty"
            GoTo NextRow

        End If

        expectedQty = CDbl(ws.Cells(r, "E").Value)


        If ws.Cells(r, "F").Value <> "" _
            And IsDate(ws.Cells(r, "F").Value) Then

            newDate = _
                Format(ws.Cells(r, "F").Value, "dd.MM.yyyy")

        Else

            ws.Cells(r, "K").Value = "Invalid Date"
            GoTo NextRow

        End If


        ws.Cells(r, "K").Value = "Processing..."

        On Error GoTo RowFail


        ' =====================================================================
        ' 1. VA02 - Locate order using Customer PO
        ' =====================================================================

        With Sess

            .findById("wnd[0]/tbar[0]/okcd").Text = "/nVA02"
            .findById("wnd[0]").sendVKey 0

            .findById("wnd[0]/usr/txtRV45S-BSTNK").Text = po
            .findById("wnd[0]/usr/btnBT_SUCH").Press


            ' Handle optional SAP dialogs

            On Error Resume Next

            .findById("wnd[1]/tbar[0]/btn[0]").Press
            .findById("wnd[1]/tbar[0]/btn[0]").Press

            Err.Clear
            On Error GoTo RowFail


            ' =================================================================
            ' 2. Position by Line Item
            ' =================================================================

            .findById( _
                "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/" & _
                "ssubSUBSCREEN_BODY:SAPMV45A:4400/" & _
                "subSUBSCREEN_TC:SAPMV45A:4900/" & _
                "subSUBSCREEN_BUTTONS:SAPMV45A:4050/" & _
                "btnBT_POPO" _
            ).Press


            .findById("wnd[1]/usr/txtRV45A-POSNR").Text = lineItem

            .findById( _
                "wnd[1]/usr/txtRV45A-POSNR" _
            ).CaretPosition = Len(lineItem)

            .findById("wnd[1]/tbar[0]/btn[0]").Press


            DoEvents

            Application.Wait _
                Now + TimeValue("0:00:01")

            .findById("wnd[0]").sendVKey 0


            ' =================================================================
            ' 3. Read SAP Part Number and Quantity
            ' =================================================================

            sapPN = _
                GetTopRowPN(.findById(gridId))

            sapQtyRaw = _
                GetTopRowQty(.findById(gridId))

            sapQty = _
                NormalizeQtyToDouble(sapQtyRaw)

        End With


        ' ---------------------------------------------------------------------
        ' Write SAP values to Excel
        ' ---------------------------------------------------------------------

        ws.Cells(r, "H").Value = sapPN
        ws.Cells(r, "I").Value = sapQty


        ' =====================================================================
        ' 4. Validate Part Number and Quantity
        ' =====================================================================

        sapPNNorm = NormalizePN(sapPN)
        excelPNNorm = NormalizePN(expectedPN)


        ' Exact comparison after normalization

        pnMatch = _
            (StrComp( _
                sapPNNorm, _
                excelPNNorm, _
                vbTextCompare _
            ) = 0)


        qtyMatch = _
            (Abs(sapQty - expectedQty) < 0.0001)


        ' ---------------------------------------------------------------------
        ' Stop update when validation fails
        ' ---------------------------------------------------------------------

        If Not pnMatch And Not qtyMatch Then

            ws.Cells(r, "K").Value = _
                "Part No. & Qty mismatch"

            GoTo NextRow


        ElseIf Not pnMatch Then

            ws.Cells(r, "K").Value = _
                "Part No. mismatch"

            GoTo NextRow


        ElseIf Not qtyMatch Then

            ws.Cells(r, "K").Value = _
                "Qty mismatch"

            GoTo NextRow

        End If


        ' =====================================================================
        ' 5. Validation passed - update delivery date
        ' =====================================================================

        cellId = _
            gridId & "/ctxtRV45A-ETDAT[6,0]"


        On Error Resume Next

        Sess.findById(cellId).Text = newDate


        If Err.Number <> 0 Then

            ws.Cells(r, "K").Value = _
                "Date update failed"

            Err.Clear

            On Error GoTo RowFail
            GoTo NextRow

        End If


        Sess.findById("wnd[0]").sendVKey 0

        On Error GoTo RowFail


        ' =====================================================================
        ' 6. Save
        ' =====================================================================

        Sess.findById("wnd[0]").sendVKey 11


        ' Handle optional confirmation dialog

        On Error Resume Next

        Sess.findById( _
            "wnd[1]/tbar[0]/btn[7]" _
        ).Press

        Err.Clear
        On Error GoTo RowFail


        ws.Cells(r, "K").Value = "Updated"

        GoTo NextRow


RowFail:

        ws.Cells(r, "K").Value = "Error"

        Err.Clear
        On Error GoTo 0


NextRow:

    Next r


    MsgBox _
        "Delivery date automation by Line Item completed.", _
        vbInformation

    Exit Sub


SAPConnectionError:

    MsgBox _
        "Unable to connect to an active SAP GUI session." & vbCrLf & _
        "Please ensure SAP GUI is open and scripting is enabled.", _
        vbExclamation

End Sub


' =============================================================================
' Read Part Number from top SAP Item Overview row
' =============================================================================

Private Function GetTopRowPN(ByVal grid As Object) As String

    Dim textVal As String

    On Error Resume Next


    textVal = _
        grid.Parent.findById( _
            grid.ID & "/ctxtRV45A-MATNR[1,0]" _
        ).Text


    If Err.Number <> 0 Or textVal = "" Then

        Err.Clear

        textVal = _
            grid.Parent.findById( _
                grid.ID & "/ctxtVBAP-MATNR[1,0]" _
            ).Text

    End If


    If Err.Number <> 0 Or textVal = "" Then

        Err.Clear

        textVal = _
            grid.Parent.findById( _
                grid.ID & "/txtRV45A-MATNR[1,0]" _
            ).Text

    End If


    On Error GoTo 0

    GetTopRowPN = Trim(textVal)

End Function


' =============================================================================
' Read Quantity from top SAP Item Overview row
' =============================================================================

Private Function GetTopRowQty(ByVal grid As Object) As String

    Dim textVal As String

    On Error Resume Next


    textVal = _
        grid.Parent.findById( _
            grid.ID & "/txtRV45A-KWMENG[2,0]" _
        ).Text


    If Err.Number <> 0 Or textVal = "" Then

        Err.Clear

        textVal = _
            grid.Parent.findById( _
                grid.ID & "/ctxtRV45A-KWMENG[2,0]" _
            ).Text

    End If


    If Err.Number <> 0 Or textVal = "" Then

        Err.Clear

        textVal = _
            grid.Parent.findById( _
                grid.ID & "/txtVBAP-KWMENG[2,0]" _
            ).Text

    End If


    On Error GoTo 0

    GetTopRowQty = Trim(textVal)

End Function


' =============================================================================
' Normalize Part Number
' =============================================================================

Private Function NormalizePN(ByVal s As String) As String

    Dim t As String

    t = Application.WorksheetFunction.Clean(CStr(s))
    t = Trim(t)
    t = Replace(t, " ", "")


    ' Remove leading zeros commonly displayed by SAP

    Do While Len(t) > 1 _
        And Left$(t, 1) = "0"

        t = Mid$(t, 2)

    Loop


    NormalizePN = t

End Function


' =============================================================================
' Convert SAP Quantity text to numeric value
' =============================================================================

Private Function NormalizeQtyToDouble(ByVal s As String) As Double

    Dim decSep As String
    Dim i As Long
    Dim ch As String
    Dim num As String


    decSep = _
        Application.International(xlDecimalSeparator)

    s = CStr(s)


    For i = 1 To Len(s)

        ch = Mid$(s, i, 1)


        If ch >= "0" And ch <= "9" Then

            num = num & ch


        ElseIf ch = "." _
            Or ch = "," _
            Or ch = decSep Then

            If InStr(num, decSep) = 0 Then
                num = num & decSep
            End If

        End If

    Next i


    If num = "" Or num = decSep Then

        NormalizeQtyToDouble = 0

    Else

        NormalizeQtyToDouble = CDbl(num)

    End If

End Function

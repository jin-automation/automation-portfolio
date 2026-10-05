Option Explicit

' =============================================================================
' SAP Sales Order Delivery Requested Date Change Automation
' =============================================================================
'
' Purpose:
'   Automates high-volume requested delivery date changes in SAP sales orders
'   using request information maintained in Excel.
'
' Excel Layout:
'
'   INPUT - Request Data
'   Column B : Customer PO
'   Column C : Line Item
'   Column D : Material / Part Number
'   Column E : Quantity
'   Column F : Requested Delivery Date
'
'   AUTOMATION OUTPUT - SAP Validation & Results
'   Column H : SAP Sales Order (Found)
'   Column I : SAP Material / Part Number
'   Column J : SAP Quantity
'   Column K : Validation
'   Column L : Status
'
' Key Control:
'   Validate first, update second.
'
'   The requested delivery date is updated only when both the Material Number
'   and Order Quantity retrieved from SAP match the expected Excel values.
'
' Technology:
'   SAP ERP / ECC
'   SAP GUI Scripting
'   Microsoft Excel
'   VBA
'
' Portfolio Note:
'   This is a sanitized portfolio version. No customer-specific, order-specific,
'   pricing, system, user, or production data is included.
'
' =============================================================================


Public Sub UpdateRequestedDeliveryDates()

    Dim SapGuiAuto As Object
    Dim SapApplication As Object
    Dim SapConnection As Object
    Dim SapSession As Object

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long

    Dim customerPO As String
    Dim lineItem As String
    Dim expectedMaterial As String
    Dim expectedQty As Double
    Dim requestedDate As String

    Dim sapSalesOrder As String
    Dim sapMaterial As String
    Dim sapQtyRaw As String
    Dim sapQty As Double

    Dim normalizedSapMaterial As String
    Dim normalizedExcelMaterial As String

    Dim materialMatches As Boolean
    Dim quantityMatches As Boolean

    Dim itemOverviewId As String
    Dim requestedDateCellId As String


    ' -------------------------------------------------------------------------
    ' SAP Item Overview table
    '
    ' SAP GUI element IDs may vary depending on SAP version, configuration,
    ' screen layout and organizational environment.
    ' -------------------------------------------------------------------------

    itemOverviewId = _
        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/" & _
        "ssubSUBSCREEN_BODY:SAPMV45A:4400/" & _
        "subSUBSCREEN_TC:SAPMV45A:4900/" & _
        "tblSAPMV45ATCTRL_U_ERF_AUFTRAG"


    ' -------------------------------------------------------------------------
    ' Connect to active SAP GUI session
    ' -------------------------------------------------------------------------

    On Error GoTo SAPConnectionError

    Set SapGuiAuto = GetObject("SAPGUI")
    Set SapApplication = SapGuiAuto.GetScriptingEngine
    Set SapConnection = SapApplication.Children(0)
    Set SapSession = SapConnection.Children(0)

    On Error GoTo 0


    ' -------------------------------------------------------------------------
    ' Excel worksheet
    ' -------------------------------------------------------------------------

    Set ws = ThisWorkbook.Sheets(1)

    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row


    ' -------------------------------------------------------------------------
    ' Process each request
    ' -------------------------------------------------------------------------

    For currentRow = 2 To lastRow

        customerPO = Trim(CStr(ws.Cells(currentRow, "B").Value))
        lineItem = Trim(CStr(ws.Cells(currentRow, "C").Value))
        expectedMaterial = Trim(CStr(ws.Cells(currentRow, "D").Value))


        ' ---------------------------------------------------------------------
        ' Read expected quantity
        ' ---------------------------------------------------------------------

        expectedQty = 0

        If ws.Cells(currentRow, "E").Value <> "" Then

            If IsNumeric(ws.Cells(currentRow, "E").Value) Then
                expectedQty = CDbl(ws.Cells(currentRow, "E").Value)
            End If

        End If


        ' ---------------------------------------------------------------------
        ' Read requested delivery date
        ' ---------------------------------------------------------------------

        If ws.Cells(currentRow, "F").Value <> "" _
            And IsDate(ws.Cells(currentRow, "F").Value) Then

            requestedDate = _
                Format(ws.Cells(currentRow, "F").Value, "dd.MM.yyyy")

        Else

            requestedDate = ""

        End If


        ' ---------------------------------------------------------------------
        ' Clear previous automation output
        '
        ' H = SAP Sales Order
        ' I = SAP Material
        ' J = SAP Qty
        ' K = Validation
        ' L = Status
        ' ---------------------------------------------------------------------

        ws.Cells(currentRow, "H").Value = ""
        ws.Cells(currentRow, "I").Value = ""
        ws.Cells(currentRow, "J").Value = ""
        ws.Cells(currentRow, "K").Value = ""
        ws.Cells(currentRow, "L").Value = ""


        ' ---------------------------------------------------------------------
        ' Validate required Excel input
        ' ---------------------------------------------------------------------

        If customerPO = "" _
            Or lineItem = "" _
            Or expectedMaterial = "" _
            Or ws.Cells(currentRow, "E").Value = "" Then

            ws.Cells(currentRow, "K").Value = "Not Processed"
            ws.Cells(currentRow, "L").Value = "SKIP: Missing data"

            GoTo NextRequest

        End If


        If requestedDate = "" Then

            ws.Cells(currentRow, "K").Value = "Not Processed"
            ws.Cells(currentRow, "L").Value = "SKIP: Invalid date"

            GoTo NextRequest

        End If


        ws.Cells(currentRow, "L").Value = "RUNNING..."

        On Error GoTo RequestError


        ' =====================================================================
        ' STEP 1
        ' Open VA02 and search using Customer PO
        ' =====================================================================

        With SapSession

            .findById("wnd[0]/tbar[0]/okcd").Text = "/nVA02"
            .findById("wnd[0]").sendVKey 0

            .findById("wnd[0]/usr/txtRV45S-BSTNK").Text = customerPO
            .findById("wnd[0]/usr/btnBT_SUCH").Press


            ' -----------------------------------------------------------------
            ' Handle optional SAP confirmation / selection dialogs
            ' -----------------------------------------------------------------

            On Error Resume Next

            .findById("wnd[1]/tbar[0]/btn[0]").Press
            .findById("wnd[1]/tbar[0]/btn[0]").Press

            On Error GoTo RequestError

            DoEvents

            Application.Wait Now + TimeValue("0:00:01")


            ' =================================================================
            ' STEP 2
            ' Capture the Sales Order found by SAP
            ' =================================================================

            sapSalesOrder = GetCurrentSalesOrder(SapSession)

            ws.Cells(currentRow, "H").Value = sapSalesOrder


            ' =================================================================
            ' STEP 3
            ' Navigate to requested line item
            ' =================================================================

            .findById( _
                "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/" & _
                "ssubSUBSCREEN_BODY:SAPMV45A:4400/" & _
                "subSUBSCREEN_TC:SAPMV45A:4900/" & _
                "subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POPO" _
            ).Press


            .findById("wnd[1]/usr/txtRV45A-POSNR").Text = lineItem

            .findById("wnd[1]/usr/txtRV45A-POSNR").CaretPosition = _
                Len(lineItem)

            .findById("wnd[1]/tbar[0]/btn[0]").Press


            DoEvents

            Application.Wait Now + TimeValue("0:00:01")

            .findById("wnd[0]").sendVKey 0


            ' =================================================================
            ' STEP 4
            ' Read SAP Material Number and Order Quantity
            ' =================================================================

            sapMaterial = _
                GetTopRowMaterial(.findById(itemOverviewId))

            sapQtyRaw = _
                GetTopRowQuantity(.findById(itemOverviewId))

            sapQty = _
                NormalizeQuantityToDouble(sapQtyRaw)

        End With


        ' ---------------------------------------------------------------------
        ' Record SAP values in Automation Output
        ' ---------------------------------------------------------------------

        ws.Cells(currentRow, "I").Value = sapMaterial
        ws.Cells(currentRow, "J").Value = sapQty


        ' =====================================================================
        ' STEP 5
        ' Validate Material Number and Quantity
        ' =====================================================================

        normalizedSapMaterial = _
            NormalizeMaterialNumber(sapMaterial)

        normalizedExcelMaterial = _
            NormalizeMaterialNumber(expectedMaterial)


        ' ---------------------------------------------------------------------
        ' Material comparison
        '
        ' SAP may display leading zeros for some material numbers.
        ' Values are normalized before comparison.
        ' ---------------------------------------------------------------------

        materialMatches = _
            (StrComp( _
                normalizedSapMaterial, _
                normalizedExcelMaterial, _
                vbTextCompare _
            ) = 0)


        ' ---------------------------------------------------------------------
        ' Quantity comparison
        ' ---------------------------------------------------------------------

        quantityMatches = _
            (Abs(sapQty - expectedQty) < 0.0001)


        ' =====================================================================
        ' STEP 6
        ' Validation decision
        ' =====================================================================

        If materialMatches And quantityMatches Then

            ws.Cells(currentRow, "K").Value = "Match"

        Else

            ws.Cells(currentRow, "K").Value = "Mismatch"


            If Not materialMatches And Not quantityMatches Then

                ws.Cells(currentRow, "L").Value = _
                    "MISMATCH: PN & Qty"

            ElseIf Not materialMatches Then

                ws.Cells(currentRow, "L").Value = _
                    "MISMATCH: PN"

            ElseIf Not quantityMatches Then

                ws.Cells(currentRow, "L").Value = _
                    "MISMATCH: Qty"

            End If


            ' -----------------------------------------------------------------
            ' SAFETY CONTROL
            '
            ' Never update SAP when Material or Quantity validation fails.
            ' -----------------------------------------------------------------

            GoTo NextRequest

        End If


        ' =====================================================================
        ' STEP 7
        ' Update requested delivery date
        ' =====================================================================

        On Error Resume Next


        requestedDateCellId = _
            itemOverviewId & "/ctxtRV45A-ETDAT[6,0]"


        SapSession.findById(requestedDateCellId).Text = _
            requestedDate


        If Err.Number <> 0 Then

            ws.Cells(currentRow, "L").Value = _
                "DATE WRITE FAIL: " & Err.Description

            Err.Clear

            GoTo NextRequest

        End If


        SapSession.findById("wnd[0]").sendVKey 0

        On Error GoTo RequestError


        ' =====================================================================
        ' STEP 8
        ' Save Sales Order
        ' =====================================================================

        SapSession.findById("wnd[0]").sendVKey 11


        ' ---------------------------------------------------------------------
        ' Handle optional SAP confirmation dialog
        ' ---------------------------------------------------------------------

        On Error Resume Next

        SapSession.findById("wnd[1]/tbar[0]/btn[7]").Press

        On Error GoTo RequestError


        ' ---------------------------------------------------------------------
        ' Successful processing
        ' ---------------------------------------------------------------------

        ws.Cells(currentRow, "L").Value = "OK"

        GoTo NextRequest


' =============================================================================
' Request-level error handling
' =============================================================================

RequestError:

        ws.Cells(currentRow, "L").Value = _
            "ERROR: " & Err.Description

        If ws.Cells(currentRow, "K").Value = "" Then
            ws.Cells(currentRow, "K").Value = "Error"
        End If

        Err.Clear

        On Error GoTo 0


NextRequest:

    Next currentRow


    MsgBox _
        "Sales Order delivery date automation completed.", _
        vbInformation

    Exit Sub


' =============================================================================
' SAP connection error
' =============================================================================

SAPConnectionError:

    MsgBox _
        "Unable to connect to an active SAP GUI session." & vbCrLf & _
        "Please ensure SAP GUI is open and scripting is enabled.", _
        vbExclamation

End Sub


' =============================================================================
' Helper Function:
' Retrieve the currently opened SAP Sales Order number
'
' IMPORTANT:
' SAP GUI element IDs may differ by system configuration.
' The function safely attempts several commonly encountered field IDs.
'
' If none of these IDs match the SAP environment, the function returns
' "(Found in SAP)" rather than generating a false Sales Order number.
' =============================================================================

Private Function GetCurrentSalesOrder( _
    ByVal SapSession As Object _
) As String

    Dim salesOrderNumber As String


    On Error Resume Next


    ' -------------------------------------------------------------------------
    ' Candidate 1
    ' Common VA02 Sales Order field
    ' -------------------------------------------------------------------------

    salesOrderNumber = _
        SapSession.findById( _
            "wnd[0]/usr/ctxtVBAK-VBELN" _
        ).Text


    ' -------------------------------------------------------------------------
    ' Candidate 2
    ' Alternative VA02 screen structure
    ' -------------------------------------------------------------------------

    If Err.Number <> 0 Or Trim(salesOrderNumber) = "" Then

        Err.Clear

        salesOrderNumber = _
            SapSession.findById( _
                "wnd[0]/usr/txtVBAK-VBELN" _
            ).Text

    End If


    ' -------------------------------------------------------------------------
    ' Candidate 3
    ' Alternative sales document field
    ' -------------------------------------------------------------------------

    If Err.Number <> 0 Or Trim(salesOrderNumber) = "" Then

        Err.Clear

        salesOrderNumber = _
            SapSession.findById( _
                "wnd[0]/usr/ctxtRV45A-VBELN" _
            ).Text

    End If


    ' -------------------------------------------------------------------------
    ' Candidate 4
    ' Alternative text field
    ' -------------------------------------------------------------------------

    If Err.Number <> 0 Or Trim(salesOrderNumber) = "" Then

        Err.Clear

        salesOrderNumber = _
            SapSession.findById( _
                "wnd[0]/usr/txtRV45A-VBELN" _
            ).Text

    End If


    On Error GoTo 0


    salesOrderNumber = Trim(salesOrderNumber)


    If salesOrderNumber = "" Then

        ' Do not invent a Sales Order number.
        ' SAP field IDs differ between environments.

        salesOrderNumber = "(Found in SAP)"

    End If


    GetCurrentSalesOrder = salesOrderNumber

End Function


' =============================================================================
' Helper Function:
' Read Material Number from first visible row of SAP Item Overview
' =============================================================================

Private Function GetTopRowMaterial( _
    ByVal grid As Object _
) As String

    Dim valueText As String


    On Error Resume Next


    valueText = _
        grid.Parent.findById( _
            grid.ID & "/ctxtRV45A-MATNR[1,0]" _
        ).Text


    If Err.Number <> 0 Or valueText = "" Then

        Err.Clear

        valueText = _
            grid.Parent.findById( _
                grid.ID & "/ctxtVBAP-MATNR[1,0]" _
            ).Text

    End If


    If Err.Number <> 0 Or valueText = "" Then

        Err.Clear

        valueText = _
            grid.Parent.findById( _
                grid.ID & "/txtRV45A-MATNR[1,0]" _
            ).Text

    End If


    On Error GoTo 0


    GetTopRowMaterial = Trim(valueText)

End Function


' =============================================================================
' Helper Function:
' Read Order Quantity from first visible row of SAP Item Overview
' =============================================================================

Private Function GetTopRowQuantity( _
    ByVal grid As Object _
) As String

    Dim valueText As String


    On Error Resume Next


    valueText = _
        grid.Parent.findById( _
            grid.ID & "/txtRV45A-KWMENG[2,0]" _
        ).Text


    If Err.Number <> 0 Or valueText = "" Then

        Err.Clear

        valueText = _
            grid.Parent.findById( _
                grid.ID & "/ctxtRV45A-KWMENG[2,0]" _
            ).Text

    End If


    If Err.Number <> 0 Or valueText = "" Then

        Err.Clear

        valueText = _
            grid.Parent.findById( _
                grid.ID & "/txtVBAP-KWMENG[2,0]" _
            ).Text

    End If


    On Error GoTo 0


    GetTopRowQuantity = Trim(valueText)

End Function


' =============================================================================
' Helper Function:
' Normalize Material / Part Number
' =============================================================================

Private Function NormalizeMaterialNumber( _
    ByVal value As String _
) As String

    Dim normalizedValue As String


    normalizedValue = _
        Application.WorksheetFunction.Clean(CStr(value))

    normalizedValue = Trim(normalizedValue)

    normalizedValue = _
        Replace(normalizedValue, " ", "")


    ' -------------------------------------------------------------------------
    ' SAP may display leading zeros for Material Numbers.
    ' Remove leading zeros before comparison.
    ' -------------------------------------------------------------------------

    Do While _
        Len(normalizedValue) > 1 _
        And Left$(normalizedValue, 1) = "0"

        normalizedValue = _
            Mid$(normalizedValue, 2)

    Loop


    NormalizeMaterialNumber = normalizedValue

End Function


' =============================================================================
' Helper Function:
' Convert SAP quantity text into numeric Double
' =============================================================================

Private Function NormalizeQuantityToDouble( _
    ByVal value As String _
) As Double

    Dim decimalSeparator As String
    Dim i As Long
    Dim currentCharacter As String
    Dim numericValue As String


    decimalSeparator = _
        Application.International(xlDecimalSeparator)

    value = CStr(value)


    For i = 1 To Len(value)

        currentCharacter = _
            Mid$(value, i, 1)


        If currentCharacter >= "0" _
            And currentCharacter <= "9" Then

            numericValue = _
                numericValue & currentCharacter


        ElseIf currentCharacter = "." _
            Or currentCharacter = "," _
            Or currentCharacter = decimalSeparator Then


            If InStr(numericValue, decimalSeparator) = 0 Then

                numericValue = _
                    numericValue & decimalSeparator

            End If

        End If

    Next i


    If numericValue = "" _
        Or numericValue = decimalSeparator Then

        NormalizeQuantityToDouble = 0

    Else

        NormalizeQuantityToDouble = _
            CDbl(numericValue)

    End If

End Function

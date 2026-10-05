Option Explicit

' =============================================================================
' SAP Sales Order Delivery Requested Date Change Automation
' =============================================================================
'
' Purpose:
'   Automates high-volume requested delivery date changes in SAP sales orders
'   using customer PO and line-item information maintained in Excel.
'
' Business Context:
'   Large reseller requests may contain more than 100 individual line items.
'   Manually locating each sales order, navigating to the requested line item,
'   validating the item, and changing the requested delivery date can require
'   significant repetitive effort.
'
' Key Controls:
'   - Searches SAP using the Customer PO number
'   - Navigates to the specified sales order line item
'   - Validates Material / Part Number before making any change
'   - Validates Order Quantity before making any change
'   - Updates the requested delivery date only after successful validation
'   - Skips mismatched or incomplete records
'   - Logs processing results and exceptions back to Excel
'
' Excel Input / Output:
'
'   Column B : Customer PO Number
'   Column C : Line Item
'   Column D : Expected Material / Part Number
'   Column E : Expected Quantity
'   Column F : Requested Delivery Date
'   Column G : Material / Quantity Match (Yes / No)
'   Column H : Processing Status
'   Column I : SAP Material (log)
'   Column J : SAP Quantity - Raw Value (log)
'   Column K : SAP Material - Normalized (log)
'   Column L : Excel Material - Normalized (log)
'   Column M : SAP Quantity - Numeric (log)
'   Column N : Excel Quantity - Numeric (log)
'
' Technology:
'   SAP ERP / ECC
'   SAP GUI Scripting
'   Microsoft Excel
'   VBA
'
' Portfolio Note:
'   This is a sanitized portfolio version.
'   No customer-specific, order-specific, pricing, or production data is
'   included in this source code.
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
    Dim requestedDate As String
    Dim expectedQty As Double

    Dim itemOverviewId As String
    Dim sapMaterial As String
    Dim sapQtyRaw As String
    Dim sapQty As Double

    Dim normalizedSapMaterial As String
    Dim normalizedExcelMaterial As String

    Dim materialMatches As Boolean
    Dim quantityMatches As Boolean

    Dim requestedDateCellId As String


    ' -------------------------------------------------------------------------
    ' SAP Item Overview table
    '
    ' Note:
    ' SAP GUI element IDs can vary depending on SAP version, configuration,
    ' screen layout, and organizational environment.
    ' -------------------------------------------------------------------------

    itemOverviewId = _
        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/" & _
        "ssubSUBSCREEN_BODY:SAPMV45A:4400/" & _
        "subSUBSCREEN_TC:SAPMV45A:4900/" & _
        "tblSAPMV45ATCTRL_U_ERF_AUFTRAG"


    ' -------------------------------------------------------------------------
    ' Connect to an active SAP GUI session
    ' -------------------------------------------------------------------------

    On Error GoTo SAPConnectionError

    Set SapGuiAuto = GetObject("SAPGUI")
    Set SapApplication = SapGuiAuto.GetScriptingEngine
    Set SapConnection = SapApplication.Children(0)
    Set SapSession = SapConnection.Children(0)

    On Error GoTo 0


    ' -------------------------------------------------------------------------
    ' Identify the Excel worksheet and processing range
    ' -------------------------------------------------------------------------

    Set ws = ThisWorkbook.Sheets(1)

    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row


    ' -------------------------------------------------------------------------
    ' Process each Excel request
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
            expectedQty = CDbl(ws.Cells(currentRow, "E").Value)
        End If


        ' ---------------------------------------------------------------------
        ' Read requested delivery date
        ' ---------------------------------------------------------------------

        If ws.Cells(currentRow, "F").Value <> "" Then
            requestedDate = Format(ws.Cells(currentRow, "F").Value, "dd.MM.yyyy")
        Else
            requestedDate = ""
        End If


        ' ---------------------------------------------------------------------
        ' Clear previous processing results
        ' ---------------------------------------------------------------------

        ws.Cells(currentRow, "G").Value = ""
        ws.Cells(currentRow, "H").Value = ""
        ws.Cells(currentRow, "I").Value = ""
        ws.Cells(currentRow, "J").Value = ""
        ws.Cells(currentRow, "K").Value = ""
        ws.Cells(currentRow, "L").Value = ""
        ws.Cells(currentRow, "M").Value = ""
        ws.Cells(currentRow, "N").Value = ""


        ' ---------------------------------------------------------------------
        ' Validate required Excel input
        ' ---------------------------------------------------------------------

        If customerPO = "" _
            Or lineItem = "" _
            Or expectedMaterial = "" _
            Or ws.Cells(currentRow, "E").Value = "" Then

            ws.Cells(currentRow, "H").Value = "SKIP: Missing data"

            GoTo NextRequest

        End If


        ws.Cells(currentRow, "H").Value = "RUNNING..."

        On Error GoTo RequestError


        ' =====================================================================
        ' STEP 1
        ' Open SAP VA02 and locate the sales order using Customer PO
        ' =====================================================================

        With SapSession

            .findById("wnd[0]/tbar[0]/okcd").Text = "/nVA02"
            .findById("wnd[0]").sendVKey 0

            .findById("wnd[0]/usr/txtRV45S-BSTNK").Text = customerPO
            .findById("wnd[0]/usr/btnBT_SUCH").Press


            ' Handle optional SAP confirmation / selection dialogs.

            On Error Resume Next

            .findById("wnd[1]/tbar[0]/btn[0]").Press
            .findById("wnd[1]/tbar[0]/btn[0]").Press

            On Error GoTo RequestError


            ' =================================================================
            ' STEP 2
            ' Navigate to the requested line item
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
            ' STEP 3
            ' Read Material Number and Order Quantity from SAP
            ' =================================================================

            sapMaterial = GetTopRowMaterial(.findById(itemOverviewId))

            sapQtyRaw = GetTopRowQuantity(.findById(itemOverviewId))

            sapQty = NormalizeQuantityToDouble(sapQtyRaw)

        End With


        ' ---------------------------------------------------------------------
        ' Record SAP values for audit / troubleshooting
        ' ---------------------------------------------------------------------

        ws.Cells(currentRow, "I").Value = sapMaterial
        ws.Cells(currentRow, "J").Value = sapQtyRaw
        ws.Cells(currentRow, "M").Value = sapQty
        ws.Cells(currentRow, "N").Value = expectedQty


        ' =====================================================================
        ' STEP 4
        ' Validate Material Number and Quantity
        ' =====================================================================

        normalizedSapMaterial = NormalizeMaterialNumber(sapMaterial)

        normalizedExcelMaterial = _
            NormalizeMaterialNumber(expectedMaterial)


        ws.Cells(currentRow, "K").Value = normalizedSapMaterial

        ws.Cells(currentRow, "L").Value = _
            normalizedExcelMaterial


        ' SAP may display leading zeros for some material numbers.
        ' The normalized values are therefore compared after removing
        ' leading zeros and spaces.

        materialMatches = _
            (InStr( _
                1, _
                normalizedSapMaterial, _
                normalizedExcelMaterial, _
                vbTextCompare _
            ) > 0) _
            Or _
            (InStr( _
                1, _
                normalizedExcelMaterial, _
                normalizedSapMaterial, _
                vbTextCompare _
            ) > 0)


        ' Quantity is compared numerically.

        quantityMatches = _
            (Abs(sapQty - expectedQty) < 0.0001)


        ' ---------------------------------------------------------------------
        ' Stop processing if validation fails
        ' ---------------------------------------------------------------------

        If materialMatches And quantityMatches Then

            ws.Cells(currentRow, "G").Value = "Yes"

        Else

            ws.Cells(currentRow, "G").Value = "No"


            If Not materialMatches And Not quantityMatches Then

                ws.Cells(currentRow, "H").Value = _
                    "MISMATCH: PN & Qty"

            ElseIf Not materialMatches Then

                ws.Cells(currentRow, "H").Value = _
                    "MISMATCH: PN"

            ElseIf Not quantityMatches Then

                ws.Cells(currentRow, "H").Value = _
                    "MISMATCH: Qty"

            End If


            ' Safety control:
            ' Never update SAP when Material or Quantity validation fails.

            GoTo NextRequest

        End If


        ' =====================================================================
        ' STEP 5
        ' Update the requested delivery date
        ' =====================================================================

        If requestedDate = "" Then

            ws.Cells(currentRow, "H").Value = "NO DATE"

            GoTo NextRequest

        End If


        On Error Resume Next


        requestedDateCellId = _
            itemOverviewId & "/ctxtRV45A-ETDAT[6,0]"


        SapSession.findById(requestedDateCellId).Text = _
            requestedDate


        If Err.Number <> 0 Then

            ws.Cells(currentRow, "H").Value = _
                "DATE WRITE FAIL: " & Err.Description

            Err.Clear

            GoTo NextRequest

        End If


        SapSession.findById("wnd[0]").sendVKey 0

        On Error GoTo RequestError


        ' =====================================================================
        ' STEP 6
        ' Save the SAP Sales Order
        ' =====================================================================

        SapSession.findById("wnd[0]").sendVKey 11


        ' Handle optional SAP confirmation dialog.

        On Error Resume Next

        SapSession.findById("wnd[1]/tbar[0]/btn[7]").Press

        On Error GoTo RequestError


        ' ---------------------------------------------------------------------
        ' Record successful processing
        ' ---------------------------------------------------------------------

        ws.Cells(currentRow, "H").Value = "OK"

        GoTo NextRequest


' =============================================================================
' Request-level error handling
' =============================================================================

RequestError:

        ws.Cells(currentRow, "H").Value = _
            "ERROR: " & Err.Description

        ws.Cells(currentRow, "G").Value = "No"

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
' Read Material Number from the first visible row of SAP Item Overview
' =============================================================================

Private Function GetTopRowMaterial(ByVal grid As Object) As String

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
' Read Order Quantity from the first visible row of SAP Item Overview
' =============================================================================

Private Function GetTopRowQuantity(ByVal grid As Object) As String

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
' Normalize Material / Part Number for comparison
' =============================================================================

Private Function NormalizeMaterialNumber(ByVal value As String) As String

    Dim normalizedValue As String


    normalizedValue = _
        Application.WorksheetFunction.Clean(CStr(value))


    normalizedValue = Trim(normalizedValue)

    normalizedValue = Replace(normalizedValue, " ", "")


    ' SAP may display leading zeros for material numbers.
    ' Remove them before comparison.

    Do While _
        Left$(normalizedValue, 1) = "0" _
        And Len(normalizedValue) > 1

        normalizedValue = Mid$(normalizedValue, 2)

    Loop


    NormalizeMaterialNumber = normalizedValue

End Function


' =============================================================================
' Helper Function:
' Convert SAP quantity text into a numeric Double value
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

        currentCharacter = Mid$(value, i, 1)


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

        NormalizeQuantityToDouble = CDbl(numericValue)

    End If

End Function

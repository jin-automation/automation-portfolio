Option Explicit

' =============================================================================
' Sales Order Delivery Requested Date Change Automation
' =============================================================================
'
' PURPOSE
' -------
' Automates high-volume requested delivery date changes in SAP ERP/ECC
' using Excel VBA and SAP GUI Scripting.
'
' The automation:
'   1. Reads request data from Excel
'   2. Searches SAP using the Customer PO
'   3. Navigates to the requested Sales Order line item
'   4. Reads the SAP Material Number and Order Quantity
'   5. Validates both values against the Excel request
'   6. Updates the requested delivery date only when validation succeeds
'   7. Saves the Sales Order
'   8. Records the validation result and processing status in Excel
'
'
' EXCEL DEMO LAYOUT
' -----------------
'
' INPUT - Request Data
'
'   B : Customer PO
'   C : Line Item
'   D : Material / Part Number
'   E : Qty
'   F : Requested Delivery Date
'
'   G : Blank separator
'
' AUTOMATION OUTPUT - SAP Validation & Results
'
'   H : SAP Sales Order (reference field in demo workbook)
'   I : SAP Material / Part Number
'   J : SAP Qty
'   K : Validation
'   L : Status
'
' Row 4 is the first data row in the public demonstration workbook.
'
'
' KEY SAFETY CONTROL
' ------------------
'
'   VALIDATE FIRST, UPDATE SECOND.
'
' The requested delivery date is changed only when both the SAP Material
' Number and Order Quantity match the expected values in Excel.
'
' If either validation fails, the SAP update is skipped and the exception
' is logged for manual review.
'
'
' PORTFOLIO NOTE
' --------------
'
' This is a sanitized portfolio version.
'
' Company-specific, customer-specific, order-specific, pricing and
' production information has been removed or generalized.
'
' SAP GUI element IDs may vary between SAP environments.
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
    Dim requestedDateCellId As String

    Dim sapMaterial As String
    Dim sapQtyRaw As String
    Dim sapQty As Double

    Dim normalizedSapMaterial As String
    Dim normalizedExcelMaterial As String

    Dim materialMatches As Boolean
    Dim quantityMatches As Boolean


    ' =========================================================================
    ' SAP Item Overview table
    ' =========================================================================

    itemOverviewId = _
        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/" & _
        "ssubSUBSCREEN_BODY:SAPMV45A:4400/" & _
        "subSUBSCREEN_TC:SAPMV45A:4900/" & _
        "tblSAPMV45ATCTRL_U_ERF_AUFTRAG"


    ' =========================================================================
    ' Connect to active SAP GUI session
    ' =========================================================================

    On Error GoTo SAPConnectionError

    Set SapGuiAuto = GetObject("SAPGUI")
    Set SapApplication = SapGuiAuto.GetScriptingEngine
    Set SapConnection = SapApplication.Children(0)
    Set SapSession = SapConnection.Children(0)

    On Error GoTo 0


    ' =========================================================================
    ' Excel worksheet
    ' =========================================================================

    Set ws = ThisWorkbook.Sheets(1)

    lastRow = _
        ws.Cells(ws.Rows.Count, "B").End(xlUp).Row


    ' =========================================================================
    ' Process each request
    '
    ' Row 2 = Section headings
    ' Row 3 = Column headings
    ' Row 4 = First data row
    ' =========================================================================

    For currentRow = 4 To lastRow


        ' ---------------------------------------------------------------------
        ' Read Excel request
        ' ---------------------------------------------------------------------

        customerPO = _
            Trim(CStr(ws.Cells(currentRow, "B").Value))

        lineItem = _
            Trim(CStr(ws.Cells(currentRow, "C").Value))

        expectedMaterial = _
            Trim(CStr(ws.Cells(currentRow, "D").Value))


        ' ---------------------------------------------------------------------
        ' Read expected quantity
        ' ---------------------------------------------------------------------

        expectedQty = 0

        If ws.Cells(currentRow, "E").Value <> "" Then

            If IsNumeric(ws.Cells(currentRow, "E").Value) Then

                expectedQty = _
                    CDbl(ws.Cells(currentRow, "E").Value)

            End If

        End If


        ' ---------------------------------------------------------------------
        ' Read requested delivery date
        ' ---------------------------------------------------------------------

        If ws.Cells(currentRow, "F").Value <> "" _
            And IsDate(ws.Cells(currentRow, "F").Value) Then

            requestedDate = _
                Format( _
                    ws.Cells(currentRow, "F").Value, _
                    "dd.MM.yyyy" _
                )

        Else

            requestedDate = ""

        End If


        ' ---------------------------------------------------------------------
        ' Clear previous automation results
        '
        ' Column H is intentionally not modified.
        ' ---------------------------------------------------------------------

        ws.Cells(currentRow, "I").Value = ""
        ws.Cells(currentRow, "J").Value = ""
        ws.Cells(currentRow, "K").Value = ""
        ws.Cells(currentRow, "L").Value = ""


        ' ---------------------------------------------------------------------
        ' Validate required input
        ' ---------------------------------------------------------------------

        If customerPO = "" _
            Or lineItem = "" _
            Or expectedMaterial = "" _
            Or ws.Cells(currentRow, "E").Value = "" Then

            ws.Cells(currentRow, "K").Value = _
                "Not Processed"

            ws.Cells(currentRow, "L").Value = _
                "SKIP: Missing data"

            GoTo NextRequest

        End If


        If requestedDate = "" Then

            ws.Cells(currentRow, "K").Value = _
                "Not Processed"

            ws.Cells(currentRow, "L").Value = _
                "NO DATE"

            GoTo NextRequest

        End If


        ws.Cells(currentRow, "L").Value = _
            "RUNNING..."


        On Error GoTo RequestError


        ' =====================================================================
        ' STEP 1
        ' Open VA02 and locate Sales Order using Customer PO
        ' =====================================================================

        With SapSession

            .findById( _
                "wnd[0]/tbar[0]/okcd" _
            ).Text = "/nVA02"

            .findById("wnd[0]").sendVKey 0


            .findById( _
                "wnd[0]/usr/txtRV45S-BSTNK" _
            ).Text = customerPO


            .findById( _
                "wnd[0]/usr/btnBT_SUCH" _
            ).Press


            ' Handle optional SAP selection / confirmation dialogs.

            On Error Resume Next

            .findById( _
                "wnd[1]/tbar[0]/btn[0]" _
            ).Press

            .findById( _
                "wnd[1]/tbar[0]/btn[0]" _
            ).Press

            On Error GoTo RequestError


            ' =================================================================
            ' STEP 2
            ' Navigate to requested line item
            ' =================================================================

            .findById( _
                "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/" & _
                "ssubSUBSCREEN_BODY:SAPMV45A:4400/" & _
                "subSUBSCREEN_TC:SAPMV45A:4900/" & _
                "subSUBSCREEN_BUTTONS:SAPMV45A:4050/" & _
                "btnBT_POPO" _
            ).Press


            .findById( _
                "wnd[1]/usr/txtRV45A-POSNR" _
            ).Text = lineItem


            .findById( _
                "wnd[1]/usr/txtRV45A-POSNR" _
            ).CaretPosition = Len(lineItem)


            .findById( _
                "wnd[1]/tbar[0]/btn[0]" _
            ).Press


            DoEvents

            Application.Wait _
                Now + TimeValue("0:00:01")


            .findById("wnd[0]").sendVKey 0


            ' =================================================================
            ' STEP 3
            ' Read Material Number and Order Quantity from SAP
            ' =================================================================

            sapMaterial = _
                GetTopRowMaterial( _
                    .findById(itemOverviewId) _
                )


            sapQtyRaw = _
                GetTopRowQuantity( _
                    .findById(itemOverviewId) _
                )


            sapQty = _
                NormalizeQuantityToDouble( _
                    sapQtyRaw _
                )

        End With


        ' =====================================================================
        ' Record SAP values in Excel
        ' =====================================================================

        ws.Cells(currentRow, "I").Value = _
            sapMaterial

        ws.Cells(currentRow, "J").Value = _
            sapQty


        ' =====================================================================
        ' STEP 4
        ' Normalize Material Numbers
        ' =====================================================================

        normalizedSapMaterial = _
            NormalizeMaterialNumber( _
                sapMaterial _
            )


        normalizedExcelMaterial = _
            NormalizeMaterialNumber( _
                expectedMaterial _
            )


        ' ---------------------------------------------------------------------
        ' Material validation
        '
        ' Leading zeros and spaces are removed before comparison.
        '
        ' The original implementation also allowed an included-within
        ' comparison to accommodate SAP Material Number formatting.
        ' ---------------------------------------------------------------------

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


        ' =====================================================================
        ' STEP 5
        ' Validate quantity
        ' =====================================================================

        quantityMatches = _
            (Abs(sapQty - expectedQty) < 0.0001)


        ' =====================================================================
        ' STEP 6
        ' Validation decision
        ' =====================================================================

        If materialMatches And quantityMatches Then

            ws.Cells(currentRow, "K").Value = _
                "Match"

        Else

            ws.Cells(currentRow, "K").Value = _
                "Mismatch"


            If Not materialMatches _
                And Not quantityMatches Then

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
        ' Update Requested Delivery Date
        ' =====================================================================

        On Error Resume Next


        requestedDateCellId = _
            itemOverviewId & _
            "/ctxtRV45A-ETDAT[6,0]"


        SapSession.findById( _
            requestedDateCellId _
        ).Text = requestedDate


        If Err.Number <> 0 Then

            ws.Cells(currentRow, "L").Value = _
                "DATE WRITE FAIL: " & _
                Err.Description

            Err.Clear

            GoTo NextRequest

        End If


        SapSession.findById( _
            "wnd[0]" _
        ).sendVKey 0


        On Error GoTo RequestError


        ' =====================================================================
        ' STEP 8
        ' Save Sales Order
        ' =====================================================================

        SapSession.findById( _
            "wnd[0]" _
        ).sendVKey 11


        ' Handle optional SAP confirmation dialog.

        On Error Resume Next

        SapSession.findById( _
            "wnd[1]/tbar[0]/btn[7]" _
        ).Press

        On Error GoTo RequestError


        ' ---------------------------------------------------------------------
        ' Successful processing
        ' ---------------------------------------------------------------------

        ws.Cells(currentRow, "L").Value = _
            "OK"

        GoTo NextRequest


' =============================================================================
' Request-level error handling
' =============================================================================

RequestError:

        ws.Cells(currentRow, "L").Value = _
            "ERROR: " & Err.Description


        If ws.Cells(currentRow, "K").Value = "" Then

            ws.Cells(currentRow, "K").Value = _
                "Error"

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
' Helper:
' Read Material Number from first visible row of SAP Item Overview
' =============================================================================

Private Function GetTopRowMaterial( _
    ByVal grid As Object _
) As String

    Dim valueText As String

    On Error Resume Next


    valueText = _
        grid.Parent.findById( _
            grid.ID & _
            "/ctxtRV45A-MATNR[1,0]" _
        ).Text


    If Err.Number <> 0 _
        Or valueText = "" Then

        Err.Clear

        valueText = _
            grid.Parent.findById( _
                grid.ID & _
                "/ctxtVBAP-MATNR[1,0]" _
            ).Text

    End If


    If Err.Number <> 0 _
        Or valueText = "" Then

        Err.Clear

        valueText = _
            grid.Parent.findById( _
                grid.ID & _
                "/txtRV45A-MATNR[1,0]" _
            ).Text

    End If


    On Error GoTo 0


    GetTopRowMaterial = _
        Trim(valueText)

End Function


' =============================================================================
' Helper:
' Read Order Quantity from first visible row of SAP Item Overview
' =============================================================================

Private Function GetTopRowQuantity( _
    ByVal grid As Object _
) As String

    Dim valueText As String

    On Error Resume Next


    valueText = _
        grid.Parent.findById( _
            grid.ID & _
            "/txtRV45A-KWMENG[2,0]" _
        ).Text


    If Err.Number <> 0 _
        Or valueText = "" Then

        Err.Clear

        valueText = _
            grid.Parent.findById( _
                grid.ID & _
                "/ctxtRV45A-KWMENG[2,0]" _
            ).Text

    End If


    If Err.Number <> 0 _
        Or valueText = "" Then

        Err.Clear

        valueText = _
            grid.Parent.findById( _
                grid.ID & _
                "/txtVBAP-KWMENG[2,0]" _
            ).Text

    End If


    On Error GoTo 0


    GetTopRowQuantity = _
        Trim(valueText)

End Function


' =============================================================================
' Helper:
' Normalize Material / Part Number
' =============================================================================

Private Function NormalizeMaterialNumber( _
    ByVal value As String _
) As String

    Dim normalizedValue As String


    normalizedValue = _
        Application.WorksheetFunction.Clean( _
            CStr(value) _
        )


    normalizedValue = _
        Trim(normalizedValue)


    normalizedValue = _
        Replace( _
            normalizedValue, _
            " ", _
            "" _
        )


    ' Remove leading zeros commonly displayed by SAP.

    Do While _
        Len(normalizedValue) > 1 _
        And Left$(normalizedValue, 1) = "0"

        normalizedValue = _
            Mid$(normalizedValue, 2)

    Loop


    NormalizeMaterialNumber = _
        normalizedValue

End Function


' =============================================================================
' Helper:
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
        Application.International( _
            xlDecimalSeparator _
        )


    value = CStr(value)


    For i = 1 To Len(value)

        currentCharacter = _
            Mid$(value, i, 1)


        If currentCharacter >= "0" _
            And currentCharacter <= "9" Then

            numericValue = _
                numericValue & _
                currentCharacter


        ElseIf currentCharacter = "." _
            Or currentCharacter = "," _
            Or currentCharacter = decimalSeparator Then


            If InStr( _
                numericValue, _
                decimalSeparator _
            ) = 0 Then

                numericValue = _
                    numericValue & _
                    decimalSeparator

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

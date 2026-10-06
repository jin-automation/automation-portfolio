Option Explicit

' =============================================================================
' SAP Delivery Date Automation
' SAP ERP/ECC | Excel VBA | SAP GUI Scripting
' =============================================================================
'
' PURPOSE
' -------
' Automates high-volume delivery date change requests in SAP ECC.
'
' WORKFLOW
' --------
' 1. Reads the delivery date change request from Excel.
' 2. Locates the requested item in SAP.
' 3. Reads the SAP Part Number and Quantity.
' 4. Validates Part Number and Quantity against the request.
' 5. Updates the delivery date only when both values match.
' 6. Logs the result in Excel.
'
' SAFETY RULE
' -----------
' Part No. and Qty must match SAP.
'
' Match    -> Delivery date is updated.
' Mismatch -> No update is made.
'
' EXCEL LAYOUT
' ------------
'
' CUSTOMER REQUEST
'   B = Customer PO
'   C = Line
'   D = Part No.
'   E = Qty
'   F = New Date
'
' SAP
'   H = SAP Part No.
'   I = SAP Qty
'
' RESULT
'   K = Status
'
' Row 4 = First request row
'
' PORTFOLIO NOTE
' --------------
' This is a sanitized portfolio version.
' Company, customer and production information has been removed or replaced.
' SAP GUI element IDs may vary between SAP environments.
'
' =============================================================================


Public Sub UpdateDeliveryDates()

    Dim SapGuiAuto As Object
    Dim SapApplication As Object
    Dim SapConnection As Object
    Dim SapSession As Object

    Dim ws As Worksheet

    Dim lastRow As Long
    Dim currentRow As Long

    Dim customerPO As String
    Dim lineItem As String
    Dim expectedPartNo As String

    Dim expectedQty As Double
    Dim requestedDate As String

    Dim sapPartNo As String
    Dim sapQtyRaw As String
    Dim sapQty As Double

    Dim normalizedSapPartNo As String
    Dim normalizedExcelPartNo As String

    Dim partMatches As Boolean
    Dim qtyMatches As Boolean

    Dim itemOverviewId As String
    Dim requestedDateCellId As String


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
    ' =========================================================================

    For currentRow = 4 To lastRow


        ' ---------------------------------------------------------------------
        ' Read customer request
        ' ---------------------------------------------------------------------

        customerPO = _
            Trim(CStr(ws.Cells(currentRow, "B").Value))

        lineItem = _
            Trim(CStr(ws.Cells(currentRow, "C").Value))

        expectedPartNo = _
            Trim(CStr(ws.Cells(currentRow, "D").Value))


        ' ---------------------------------------------------------------------
        ' Clear previous SAP results
        ' ---------------------------------------------------------------------

        ws.Cells(currentRow, "H").Value = ""
        ws.Cells(currentRow, "I").Value = ""
        ws.Cells(currentRow, "K").Value = ""


        ' ---------------------------------------------------------------------
        ' Validate required input
        ' ---------------------------------------------------------------------

        If customerPO = "" _
            Or lineItem = "" _
            Or expectedPartNo = "" _
            Or ws.Cells(currentRow, "E").Value = "" Then

            ws.Cells(currentRow, "K").Value = _
                "Missing data"

            GoTo NextRequest

        End If


        ' ---------------------------------------------------------------------
        ' Validate quantity
        ' ---------------------------------------------------------------------

        If Not IsNumeric(ws.Cells(currentRow, "E").Value) Then

            ws.Cells(currentRow, "K").Value = _
                "Invalid Qty"

            GoTo NextRequest

        End If


        expectedQty = _
            CDbl(ws.Cells(currentRow, "E").Value)


        ' ---------------------------------------------------------------------
        ' Validate requested delivery date
        ' ---------------------------------------------------------------------

        If ws.Cells(currentRow, "F").Value <> "" _
            And IsDate(ws.Cells(currentRow, "F").Value) Then

            requestedDate = _
                Format( _
                    ws.Cells(currentRow, "F").Value, _
                    "dd.MM.yyyy" _
                )

        Else

            ws.Cells(currentRow, "K").Value = _
                "Invalid Date"

            GoTo NextRequest

        End If


        ws.Cells(currentRow, "K").Value = _
            "Processing..."


        On Error GoTo RequestError


        ' =====================================================================
        ' STEP 1
        ' Open SAP Sales Order change transaction
        ' =====================================================================

        With SapSession

            .findById( _
                "wnd[0]/tbar[0]/okcd" _
            ).Text = "/nVA02"

            .findById("wnd[0]").sendVKey 0


            ' -----------------------------------------------------------------
            ' Locate request using Customer PO
            ' -----------------------------------------------------------------

            .findById( _
                "wnd[0]/usr/txtRV45S-BSTNK" _
            ).Text = customerPO


            .findById( _
                "wnd[0]/usr/btnBT_SUCH" _
            ).Press


            ' -----------------------------------------------------------------
            ' Handle optional SAP dialogs
            ' -----------------------------------------------------------------

            On Error Resume Next

            .findById( _
                "wnd[1]/tbar[0]/btn[0]" _
            ).Press

            .findById( _
                "wnd[1]/tbar[0]/btn[0]" _
            ).Press

            Err.Clear

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
            ' Read Part Number and Quantity from SAP
            ' =================================================================

            sapPartNo = _
                GetTopRowPartNo( _
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
        ' STEP 4
        ' Write SAP values to Excel
        ' =====================================================================

        ws.Cells(currentRow, "H").Value = _
            sapPartNo

        ws.Cells(currentRow, "I").Value = _
            sapQty


        ' =====================================================================
        ' STEP 5
        ' Normalize Part Numbers
        ' =====================================================================

        normalizedSapPartNo = _
            NormalizePartNumber( _
                sapPartNo _
            )


        normalizedExcelPartNo = _
            NormalizePartNumber( _
                expectedPartNo _
            )


        ' ---------------------------------------------------------------------
        ' Exact comparison after normalization
        ' ---------------------------------------------------------------------

        partMatches = _
            (StrComp( _
                normalizedSapPartNo, _
                normalizedExcelPartNo, _
                vbTextCompare _
            ) = 0)


        ' =====================================================================
        ' STEP 6
        ' Compare Quantity
        ' =====================================================================

        qtyMatches = _
            (Abs(sapQty - expectedQty) < 0.0001)


        ' =====================================================================
        ' STEP 7
        ' Validation decision
        ' =====================================================================

        If Not partMatches _
            And Not qtyMatches Then

            ws.Cells(currentRow, "K").Value = _
                "Part No. & Qty mismatch"

            GoTo NextRequest


        ElseIf Not partMatches Then

            ws.Cells(currentRow, "K").Value = _
                "Part No. mismatch"

            GoTo NextRequest


        ElseIf Not qtyMatches Then

            ws.Cells(currentRow, "K").Value = _
                "Qty mismatch"

            GoTo NextRequest

        End If


        ' =====================================================================
        ' STEP 8
        ' Part Number and Quantity match.
        ' Update delivery date.
        ' =====================================================================

        requestedDateCellId = _
            itemOverviewId & _
            "/ctxtRV45A-ETDAT[6,0]"


        On Error Resume Next


        SapSession.findById( _
            requestedDateCellId _
        ).Text = requestedDate


        If Err.Number <> 0 Then

            ws.Cells(currentRow, "K").Value = _
                "Date update failed"

            Err.Clear

            On Error GoTo RequestError

            GoTo NextRequest

        End If


        SapSession.findById( _
            "wnd[0]" _
        ).sendVKey 0


        On Error GoTo RequestError


        ' =====================================================================
        ' STEP 9
        ' Save SAP change
        ' =====================================================================

        SapSession.findById( _
            "wnd[0]" _
        ).sendVKey 11


        ' ---------------------------------------------------------------------
        ' Handle optional SAP confirmation dialog
        ' ---------------------------------------------------------------------

        On Error Resume Next

        SapSession.findById( _
            "wnd[1]/tbar[0]/btn[7]" _
        ).Press

        Err.Clear

        On Error GoTo RequestError


        ' ---------------------------------------------------------------------
        ' Successful update
        ' ---------------------------------------------------------------------

        ws.Cells(currentRow, "K").Value = _
            "Updated"


        GoTo NextRequest


' =============================================================================
' Request-level error handling
' =============================================================================

RequestError:

        ws.Cells(currentRow, "K").Value = _
            "Error"

        Err.Clear

        On Error GoTo 0


NextRequest:

    Next currentRow


    MsgBox _
        "Delivery date automation completed.", _
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
' Read Part Number from first visible row in SAP Item Overview
' =============================================================================

Private Function GetTopRowPartNo( _
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


    GetTopRowPartNo = _
        Trim(valueText)

End Function


' =============================================================================
' Helper:
' Read Quantity from first visible row in SAP Item Overview
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
' Normalize SAP / Excel Part Number before comparison
' =============================================================================

Private Function NormalizePartNumber( _
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


    NormalizePartNumber = _
        normalizedValue

End Function


' =============================================================================
' Helper:
' Convert SAP quantity text to numeric value
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

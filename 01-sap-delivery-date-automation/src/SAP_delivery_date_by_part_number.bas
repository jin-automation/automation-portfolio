
Option Explicit

' =============================================================================
' SAP Delivery Date Automation - By Part Number
' SAP ERP/ECC | Excel VBA | SAP GUI Scripting
'
' PURPOSE
' -------
' Updates requested delivery dates in SAP ECC by locating the requested
' item using its Part Number / Material.
'
' EXCEL INPUT
' -----------
' A = Customer PO
' B = Part Number / Material
' E = New Delivery Date
'
' OUTPUT
' ------
' F = Status
'
' First data row = 2
'
' Portfolio version - company/customer data removed.
' SAP GUI element IDs may vary between SAP environments.
' =============================================================================

Public Sub UpdateDeliveryDate_ByPartNumber()

    Dim Sap As Object
    Dim App As Object
    Dim Conn As Object
    Dim Sess As Object

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Dim po As String
    Dim part As String
    Dim newDate As String

    Dim gridId As String


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
        ws.Cells(ws.Rows.Count, "A").End(xlUp).Row


    ' =========================================================================
    ' Process each request
    ' =========================================================================

    For r = 2 To lastRow

        po = Trim(CStr(ws.Cells(r, "A").Value))
        part = Trim(CStr(ws.Cells(r, "B").Value))


        ' ---------------------------------------------------------------------
        ' Validate required input
        ' ---------------------------------------------------------------------

        If po = "" _
            Or part = "" Then

            ws.Cells(r, "F").Value = "SKIP"
            GoTo NextRow

        End If


        If ws.Cells(r, "E").Value <> "" _
            And IsDate(ws.Cells(r, "E").Value) Then

            newDate = _
                Format( _
                    ws.Cells(r, "E").Value, _
                    "dd.MM.yyyy" _
                )

        Else

            ws.Cells(r, "F").Value = "INVALID DATE"
            GoTo NextRow

        End If


        ws.Cells(r, "F").Value = "PROCESSING..."

        On Error GoTo RowFail


        ' =====================================================================
        ' 1. VA02 - Locate order using Customer PO
        ' =====================================================================

        Sess.findById( _
            "wnd[0]/tbar[0]/okcd" _
        ).Text = "/nVA02"

        Sess.findById("wnd[0]").sendVKey 0


        Sess.findById( _
            "wnd[0]/usr/txtRV45S-BSTNK" _
        ).Text = po


        Sess.findById( _
            "wnd[0]/usr/btnBT_SUCH" _
        ).Press


        ' ---------------------------------------------------------------------
        ' Handle optional SAP dialogs
        ' ---------------------------------------------------------------------

        On Error Resume Next

        Sess.findById( _
            "wnd[1]/tbar[0]/btn[0]" _
        ).Press

        Sess.findById( _
            "wnd[1]/tbar[0]/btn[0]" _
        ).Press

        Err.Clear
        On Error GoTo RowFail


        ' =====================================================================
        ' 2. Position using Part Number / Material
        ' =====================================================================

        Sess.findById( _
            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/" & _
            "ssubSUBSCREEN_BODY:SAPMV45A:4400/" & _
            "subSUBSCREEN_TC:SAPMV45A:4900/" & _
            "subSUBSCREEN_BUTTONS:SAPMV45A:4050/" & _
            "btnBT_POPO" _
        ).Press


        Sess.findById( _
            "wnd[1]/usr/ctxtRV45A-PO_MATNR" _
        ).Text = part


        Sess.findById( _
            "wnd[1]/tbar[0]/btn[0]" _
        ).Press


        DoEvents

        Application.Wait _
            Now + TimeValue("0:00:01")


        Sess.findById("wnd[0]").sendVKey 0


        ' =====================================================================
        ' 3. Update First Date
        '
        ' First attempt:
        ' Write to the positioned top row.
        '
        ' Fallback:
        ' Scan the item table for the requested material.
        ' =====================================================================

        If Not TrySetFirstDate( _
            Sess, _
            gridId, _
            0, _
            newDate _
        ) Then


            If Not TrySetFirstDateByScan( _
                Sess, _
                gridId, _
                part, _
                newDate _
            ) Then

                ws.Cells(r, "F").Value = _
                    "NOT FOUND / WRITE FAIL"

                GoTo NextRow

            End If

        End If


        ' =====================================================================
        ' 4. Save
        ' =====================================================================

        Sess.findById("wnd[0]").sendVKey 11


        ' ---------------------------------------------------------------------
        ' Handle optional confirmation dialog
        ' ---------------------------------------------------------------------

        On Error Resume Next

        Sess.findById( _
            "wnd[1]/tbar[0]/btn[7]" _
        ).Press

        Err.Clear
        On Error GoTo RowFail


        ws.Cells(r, "F").Value = "Updated"

        GoTo NextRow


RowFail:

        ws.Cells(r, "F").Value = "Error"

        Err.Clear
        On Error GoTo 0


NextRow:

    Next r


    MsgBox _
        "Delivery date automation by Part Number completed.", _
        vbInformation

    Exit Sub


SAPConnectionError:

    MsgBox _
        "Unable to connect to an active SAP GUI session." & vbCrLf & _
        "Please ensure SAP GUI is open and scripting is enabled.", _
        vbExclamation

End Sub


' =============================================================================
' Try to write First Date to a specified row.
'
' SAP layouts may expose the date field using different column indexes
' or control types, so several known variants are attempted.
' =============================================================================

Private Function TrySetFirstDate( _
    ByVal Sess As Object, _
    ByVal gridId As String, _
    ByVal rowIdx As Long, _
    ByVal dateStr As String _
) As Boolean

    Dim ok As Boolean

    Dim i As Long
    Dim j As Long

    Dim candId As String

    Dim colIdxArr As Variant
    Dim prefixArr As Variant


    ok = False


    ' Possible column indexes

    colIdxArr = _
        Array(5, 6, 4)


    ' Possible SAP field/control types

    prefixArr = _
        Array( _
            "ctxtRV45A-ETDAT", _
            "txtRV45A-ETDAT", _
            "ctxtRV45A-EDATU", _
            "txtRV45A-EDATU" _
        )


    On Error Resume Next


    For i = LBound(colIdxArr) _
        To UBound(colIdxArr)


        For j = LBound(prefixArr) _
            To UBound(prefixArr)


            candId = _
                gridId & "/" & _
                prefixArr(j) & _
                "[" & _
                colIdxArr(i) & _
                "," & _
                rowIdx & _
                "]"


            Err.Clear


            Sess.findById( _
                candId _
            ).Text = dateStr


            If Err.Number = 0 Then

                Sess.findById( _
                    "wnd[0]" _
                ).sendVKey 0

                ok = True

                Exit For

            End If

        Next j


        If ok Then Exit For

    Next i


    On Error GoTo 0


    TrySetFirstDate = ok

End Function


' =============================================================================
' Fallback:
' Scan SAP Item Overview for matching Part Number / Material.
' =============================================================================

Private Function TrySetFirstDateByScan( _
    ByVal Sess As Object, _
    ByVal gridId As String, _
    ByVal material As String, _
    ByVal dateStr As String _
) As Boolean

    Dim grid As Object

    Dim rows As Long
    Dim i As Long

    Dim matVal As String
    Dim ok As Boolean


    Set grid = _
        Sess.findById(gridId)


    rows = grid.RowCount

    ok = False


    For i = 0 To rows - 1

        matVal = ""


        On Error Resume Next


        matVal = _
            Trim( _
                grid.GetCellValue( _
                    i, _
                    "RV45A-MATNR" _
                ) _
            )


        If Err.Number <> 0 _
            Or matVal = "" Then

            Err.Clear

            matVal = _
                Trim( _
                    grid.GetCellValue( _
                        i, _
                        "MATNR" _
                    ) _
                )

        End If


        On Error GoTo 0


        ' ---------------------------------------------------------------------
        ' Requested material found
        ' ---------------------------------------------------------------------

        If StrComp( _
            NormalizeMaterial(matVal), _
            NormalizeMaterial(material), _
            vbTextCompare _
        ) = 0 Then

            ok = _
                TrySetFirstDate( _
                    Sess, _
                    gridId, _
                    i, _
                    dateStr _
                )

            Exit For

        End If

    Next i


    TrySetFirstDateByScan = ok

End Function


' =============================================================================
' Normalize Material / Part Number
'
' Removes spaces and leading zeros commonly displayed by SAP.
' =============================================================================

Private Function NormalizeMaterial( _
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


    Do While _
        Len(normalizedValue) > 1 _
        And Left$(normalizedValue, 1) = "0"

        normalizedValue = _
            Mid$(normalizedValue, 2)

    Loop


    NormalizeMaterial = normalizedValue

End Function

Option Explicit

' ============================================================
' SAP Customer PO Lookup & Excel Update Automation
'
' Portfolio Version
'
' Workflow:
' Excel Sales Order
'       ↓
' SAP VA03
'       ↓
' Customer PO Lookup
'       ↓
' Excel Update
'
' Demo Excel Mapping:
' A = SALES ORDER
' E = CUSTOMER PO
' G = STATUS
'
' Data starts from Row 5.
' ============================================================


' ===== USER OPTIONS =====

' False = Do not overwrite an existing Customer PO
' True  = Refresh/overwrite existing Customer PO values
Private Const DO_OVERWRITE_FILLED As Boolean = False


' ===== SAP OBJECTS =====

Private SapGuiAuto As Object
Private objGui As Object
Private objConn As Object
Private session As Object


' ============================================================
' MAIN
' ============================================================

Public Sub UpdateCustomerPO_FromSAP()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Dim salesOrder As String
    Dim currentPO As String
    Dim customerPO As String

    Set ws = ActiveSheet

    ' Connect to an already-open SAP GUI session
    If Not AttachSAP() Then
        MsgBox "Unable to connect to SAP." & vbCrLf & _
               "Please make sure SAP GUI is open and logged in.", _
               vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Sales Order is stored in Column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Demo data starts from Row 5
    For r = 5 To lastRow

        salesOrder = Trim$(CStr(ws.Cells(r, "A").Value))
        currentPO = Trim$(CStr(ws.Cells(r, "E").Value))

        ' Skip blank rows
        If salesOrder <> "" Then

            ' Do not overwrite an existing Customer PO by default
            If currentPO = "" Or DO_OVERWRITE_FILLED Then

                customerPO = GetCustomerPO_From_VA03(salesOrder)

                Select Case customerPO

                    Case "#NOT_FOUND"

                        ws.Cells(r, "E").Value = "#NOT_FOUND"
                        ws.Cells(r, "G").Value = "Order Not Found"

                    Case "#PO_NOT_FOUND"

                        ws.Cells(r, "E").Value = "#PO_NOT_FOUND"
                        ws.Cells(r, "G").Value = "PO Not Found"

                    Case "#ERROR"

                        ws.Cells(r, "E").Value = "#ERROR"
                        ws.Cells(r, "G").Value = "Error"

                    Case Else

                        ws.Cells(r, "E").Value = customerPO
                        ws.Cells(r, "G").Value = "Updated"

                End Select

                DoEvents

            Else

                ws.Cells(r, "G").Value = "Existing PO - Skipped"

            End If

        End If

    Next r

    Application.ScreenUpdating = True

    MsgBox "Customer PO lookup completed.", vbInformation

End Sub


' ============================================================
' CUSTOMER PO LOOKUP FROM SAP VA03
' ============================================================

Private Function GetCustomerPO_From_VA03( _
    ByVal salesOrder As String) As String

    On Error GoTo EH

    ' Open VA03 - Display Sales Order
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nva03"
    session.findById("wnd[0]").sendVKey 0

    ' Enter Sales Order number
    session.findById( _
        "wnd[0]/usr/ctxtVBAK-VBELN").Text = salesOrder

    session.findById("wnd[0]").sendVKey 0


    ' Check whether SAP returned an error
    If session.findById("wnd[0]/sbar").MessageType = "E" Then

        GetCustomerPO_From_VA03 = "#NOT_FOUND"
        Exit Function

    End If


    ' Possible Customer PO field locations.
    '
    ' Different SAP screen configurations may expose the
    ' Customer PO / Customer Reference field using different IDs.

    Dim fieldIDs As Variant
    Dim i As Long
    Dim value As String

    fieldIDs = Array( _
        "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD", _
        "wnd[0]/usr/subSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD", _
        "wnd[0]/usr/ctxtVBAK-BSTNK", _
        "wnd[0]/usr/txtVBAK-BSTNK", _
        "wnd[0]/usr/ctxtVBKD-BSTKD", _
        "wnd[0]/usr/txtVBKD-BSTKD" _
    )


    ' Try each possible field location
    For i = LBound(fieldIDs) To UBound(fieldIDs)

        If ExistsById(fieldIDs(i)) Then

            value = Trim$(CStr( _
                session.findById(fieldIDs(i)).Text))

            If value <> "" Then

                GetCustomerPO_From_VA03 = value
                Exit Function

            End If

        End If

    Next i


    ' Sales Order exists, but Customer PO was not found
    GetCustomerPO_From_VA03 = "#PO_NOT_FOUND"

    Exit Function


EH:

    GetCustomerPO_From_VA03 = "#ERROR"

End Function


' ============================================================
' CONNECT TO SAP GUI
' ============================================================

Private Function AttachSAP() As Boolean

    On Error GoTo EH

    Set SapGuiAuto = GetObject("SAPGUI")
    Set objGui = SapGuiAuto.GetScriptingEngine
    Set objConn = objGui.Children(0)
    Set session = objConn.Children(0)

    AttachSAP = True
    Exit Function


EH:

    AttachSAP = False

End Function


' ============================================================
' CHECK WHETHER SAP CONTROL EXISTS
' ============================================================

Private Function ExistsById( _
    ByVal fullId As String) As Boolean

    On Error GoTo NotFound

    Dim obj As Object

    Set obj = session.findById(fullId)

    ExistsById = True
    Exit Function


NotFound:

    ExistsById = False

End Function

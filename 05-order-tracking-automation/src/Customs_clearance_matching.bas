
Option Explicit

' Portfolio version
' Matches internal Sales Orders against an open Customs Clearance workbook.
' Fictional/generic worksheet and column names are used.

Private Const INTERNAL_SHEET As String = "ORDER"
Private Const INTERNAL_SO_COL As String = "B"
Private Const INTERNAL_PO_COL As String = "C"
Private Const INTERNAL_START_ROW As Long = 3

Private Const ONLINE_SO_COL As String = "A"
Private Const ONLINE_START_ROW As Long = 3

Private Const CLEARANCE_NAME_KEY As String = "CustomsClearance"
Private Const CLEARANCE_START_ROW As Long = 11
Private Const CLEARANCE_ORDER_COL As String = "G"
Private Const CLEARANCE_REMARK_COL As String = "O"
Private Const CLEARANCE_PO_COL As String = "Q"

Public Sub Match_Customs_Clearance_List()

    Dim wsOrder As Worksheet, wsOnline As Worksheet
    Dim wbClearance As Workbook, wsClearance As Worksheet
    Dim dictPO As Object, dictOnline As Object
    Dim lastRow As Long, r As Long
    Dim orderKey As String, remark As String
    Dim matchedCount As Long, poUpdatedCount As Long

    Set wsOrder = ThisWorkbook.Worksheets(INTERNAL_SHEET)
    Set wsOnline = GetSheetRightOf(wsOrder)

    Set wbClearance = FindOpenWorkbookLike(CLEARANCE_NAME_KEY)
    If wbClearance Is Nothing Then
        MsgBox "Open the Customs Clearance workbook first.", vbExclamation
        Exit Sub
    End If

    Set wsClearance = wbClearance.ActiveSheet

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo CleanFail

    Set dictPO = CreateObject("Scripting.Dictionary")
    dictPO.CompareMode = vbTextCompare

    lastRow = LastRowInColumn(wsOrder, INTERNAL_SO_COL)
    For r = INTERNAL_START_ROW To lastRow
        orderKey = NormalizeKey(wsOrder.Cells(r, INTERNAL_SO_COL).Value)
        If Len(orderKey) > 0 Then
            dictPO(orderKey) = wsOrder.Cells(r, INTERNAL_PO_COL).Value
        End If
    Next r

    Set dictOnline = CreateObject("Scripting.Dictionary")
    dictOnline.CompareMode = vbTextCompare

    If Not wsOnline Is Nothing Then
        lastRow = LastRowInColumn(wsOnline, ONLINE_SO_COL)
        For r = ONLINE_START_ROW To lastRow
            orderKey = NormalizeKey(wsOnline.Cells(r, ONLINE_SO_COL).Value)
            If Len(orderKey) > 0 Then dictOnline(orderKey) = True
        Next r
    End If

    lastRow = LastRowInColumn(wsClearance, CLEARANCE_ORDER_COL)
    ClearYellowFill wsClearance.Range(CLEARANCE_ORDER_COL & CLEARANCE_START_ROW & ":" & CLEARANCE_ORDER_COL & lastRow)

    For r = CLEARANCE_START_ROW To lastRow

        orderKey = NormalizeKey(wsClearance.Cells(r, CLEARANCE_ORDER_COL).Value)
        If Len(orderKey) = 0 Then GoTo NextRow

        If dictPO.Exists(orderKey) Or dictOnline.Exists(orderKey) Then

            wsClearance.Cells(r, CLEARANCE_ORDER_COL).Interior.Color = vbYellow
            matchedCount = matchedCount + 1

            If dictPO.Exists(orderKey) Then
                remark = CStr(wsClearance.Cells(r, CLEARANCE_REMARK_COL).Value)

                If IsEligibleForPOUpdate(remark) Then
                    wsClearance.Cells(r, CLEARANCE_PO_COL).Value = dictPO(orderKey)
                    poUpdatedCount = poUpdatedCount + 1
                End If
            End If
        End If

NextRow:
    Next r

    MsgBox "Customs Clearance matching completed." & vbCrLf & _
           "Matched orders: " & matchedCount & vbCrLf & _
           "PO values updated: " & poUpdatedCount, vbInformation

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    MsgBox "Error: " & Err.Description, vbExclamation
    Resume CleanExit

End Sub

Private Function IsEligibleForPOUpdate(ByVal remark As String) As Boolean
    ' Portfolio version: blank remarks are eligible.
    ' Replace with business-specific rules when required.
    IsEligibleForPOUpdate = (Len(Trim$(remark)) = 0)
End Function

Private Function GetSheetRightOf(ByVal ws As Worksheet) As Worksheet
    If ws.Index < ws.Parent.Worksheets.Count Then
        Set GetSheetRightOf = ws.Parent.Worksheets(ws.Index + 1)
    End If
End Function

Private Function FindOpenWorkbookLike(ByVal namePart As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If InStr(1, wb.Name, namePart, vbTextCompare) > 0 Then
            Set FindOpenWorkbookLike = wb
            Exit Function
        End If
    Next wb
End Function

Private Function LastRowInColumn(ByVal ws As Worksheet, ByVal colLetter As String) As Long
    LastRowInColumn = ws.Cells(ws.Rows.Count, colLetter).End(xlUp).Row
End Function

Private Function NormalizeKey(ByVal v As Variant) As String
    Dim s As String
    s = Trim$(CStr(v))
    s = Replace(s, " ", "")
    s = Replace(s, ",", "")
    s = Replace(s, "'", "")

    If IsAllDigits(s) Then
        Do While Len(s) > 1 And Left$(s, 1) = "0"
            s = Mid$(s, 2)
        Loop
    End If

    NormalizeKey = s
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function

    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i

    IsAllDigits = True
End Function

Private Sub ClearYellowFill(ByVal rng As Range)
    Dim c As Range
    For Each c In rng.Cells
        If c.Interior.Color = vbYellow Then c.Interior.Pattern = xlNone
    Next c
End Sub


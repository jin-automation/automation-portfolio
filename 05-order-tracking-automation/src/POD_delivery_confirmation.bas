
Option Explicit

' Portfolio version
' Matches POD Order Numbers against internal order worksheets and updates
' Delivery Status and POD Date. Fictional/generic names are used.

Private Const POD_SHEET As String = "POD"
Private Const POD_ORDER_COL As String = "A"
Private Const POD_DATE_COL As String = "B"
Private Const POD_START_ROW As Long = 2

Private Const INTERNAL_START_ROW As Long = 5
Private Const ORDER_COL As String = "A"
Private Const POD_TARGET_COL As String = "D"
Private Const STATUS_COL As String = "E"

Public Sub Match_POD_And_Update_Status()

    Dim wbPOD As Workbook, wsPOD As Worksheet
    Dim ws As Worksheet
    Dim dictPOD As Object
    Dim lastRow As Long, r As Long
    Dim orderKey As String
    Dim updatedCount As Long

    Set wbPOD = FindOpenWorkbookLike("POD")
    If wbPOD Is Nothing Then
        MsgBox "Open the POD workbook first.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wsPOD = wbPOD.Worksheets(POD_SHEET)
    On Error GoTo 0
    If wsPOD Is Nothing Then Set wsPOD = wbPOD.ActiveSheet

    Set dictPOD = CreateObject("Scripting.Dictionary")
    dictPOD.CompareMode = vbTextCompare

    lastRow = wsPOD.Cells(wsPOD.Rows.Count, POD_ORDER_COL).End(xlUp).Row

    For r = POD_START_ROW To lastRow

        orderKey = NormalizeKey(wsPOD.Cells(r, POD_ORDER_COL).Value)

        If Len(orderKey) > 0 And IsDate(wsPOD.Cells(r, POD_DATE_COL).Value) Then

            If dictPOD.Exists(orderKey) Then
                If CDate(wsPOD.Cells(r, POD_DATE_COL).Value) > CDate(dictPOD(orderKey)) Then
                    dictPOD(orderKey) = CDate(wsPOD.Cells(r, POD_DATE_COL).Value)
                End If
            Else
                dictPOD.Add orderKey, CDate(wsPOD.Cells(r, POD_DATE_COL).Value)
            End If

        End If
    Next r

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo CleanFail

    For Each ws In ThisWorkbook.Worksheets

        If ws.Name <> "Demo" Then

            lastRow = ws.Cells(ws.Rows.Count, ORDER_COL).End(xlUp).Row

            For r = INTERNAL_START_ROW To lastRow

                orderKey = NormalizeKey(ws.Cells(r, ORDER_COL).Value)

                If dictPOD.Exists(orderKey) Then
                    ws.Cells(r, POD_TARGET_COL).Value = dictPOD(orderKey)
                    ws.Cells(r, POD_TARGET_COL).NumberFormat = "dd-mmm-yy"
                    ws.Cells(r, STATUS_COL).Value = "Completed"
                    updatedCount = updatedCount + 1
                End If

            Next r
        End If

    Next ws

    MsgBox "POD matching completed." & vbCrLf & _
           "Rows updated: " & updatedCount, vbInformation

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    MsgBox "Error: " & Err.Description, vbExclamation
    Resume CleanExit

End Sub

Private Function FindOpenWorkbookLike(ByVal namePart As String) As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            If InStr(1, wb.Name, namePart, vbTextCompare) > 0 Then
                Set FindOpenWorkbookLike = wb
                Exit Function
            End If
        End If
    Next wb
End Function

Private Function NormalizeKey(ByVal v As Variant) As String
    Dim s As String
    s = Trim$(CStr(v))
    s = Replace(s, " ", "")
    s = Replace(s, ",", "")
    s = Replace(s, "'", "")
    s = Replace(s, vbTab, "")

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


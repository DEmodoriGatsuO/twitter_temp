Option Explicit

'シート名
Private Const SHEET_NAME      As String = "Sheet1"
'白色
Private Const WHITE           As Long = 16777215
Private Const INPUT_RANGE_STR As String = "$C$3:$G$7"
Private Const CHECK_RANGE_STR As String = "$C$10:$G$14"
Private Const START_NUM       As Integer = 1
Private Const Column_NUM      As Integer = 2 '妥協
Private Const RECORD_ROW_NUM  As Integer = 16
Sub kanikaniSample1()
    
    Dim INPUT_RANGE     As Range
    Dim CHECK_RANGE     As Range
    Dim rg              As Range
    Dim last_row_num    As Long
    Dim last_column_num As Long
    Dim i               As Long
    Dim j               As Long
    
    With ThisWorkbook.Sheets(SHEET_NAME)
        Set INPUT_RANGE = .Range(INPUT_RANGE_STR)
        Set CHECK_RANGE = .Range(CHECK_RANGE_STR)
        
        last_row_num = UBound(INPUT_RANGE.Value, 1)
        last_column_num = UBound(INPUT_RANGE.Value, 2)
        
        For i = START_NUM To last_row_num
            For j = START_NUM To last_column_num
                If INPUT_RANGE.Cells(i, j).Interior.Color <> _
                    CHECK_RANGE.Cells(i, j).Interior.Color Then
                        .Cells(RECORD_ROW_NUM, j + Column_NUM).Value = "不一致"
                End If
            Next j
        Next i
        
    End With
    
    Set INPUT_RANGE = Nothing
    Set CHECK_RANGE = Nothing

End Sub

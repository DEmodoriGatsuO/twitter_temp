Option Explicit

'#########################################
'## 定数                                 #
'#########################################
'シート名
Private Const SHEET_NAME      As String = "Sheet1"
'範囲文字列
Private Const INPUT_RANGE_STR As String = "$C$3:$G$7"
Private Const CHECK_RANGE_STR As String = "$C$10:$G$14"
'配列の数値ズレ用
Private Const START_NUM       As Integer = 1
'妥協
'- C列始まりなので・・・
'- 入力する行番号
Private Const Column_NUM      As Integer = 2
Private Const RECORD_ROW_NUM  As Integer = 16
Private Const INPUT_VALUE     As String = "不一致"
'#########################################
'## カニカニさんお題 20230113            #
'#########################################
Sub kanikaniSample1()
    
    'オブジェクト
    Dim INPUT_RANGE     As Range
    Dim CHECK_RANGE     As Range
    '最終行列
    Dim last_row_num    As Long
    Dim last_column_num As Long
    'イテレータ
    Dim i               As Long
    Dim j               As Long
    
    'ワークブックとシート名決め打ち
    With ThisWorkbook.Sheets(SHEET_NAME)
        'オブジェクト宣言
        Set INPUT_RANGE = .Range(INPUT_RANGE_STR)
        Set CHECK_RANGE = .Range(CHECK_RANGE_STR)
            
        '最終行列は配列の終わり値で求める
        last_row_num = UBound(INPUT_RANGE.Value, 1)
        last_column_num = UBound(INPUT_RANGE.Value, 2)
        
        'Loop
        For i = START_NUM To last_row_num
            For j = START_NUM To last_column_num
                If INPUT_RANGE.Cells(i, j).Interior.Color <> _
                    CHECK_RANGE.Cells(i, j).Interior.Color Then
                    
                        .Cells(RECORD_ROW_NUM, j + Column_NUM).Value = INPUT_VALUE
                        '差異がひとつでもあったらLoopから抜ける
                        Exit For
                End If
            Next j
        Next i
        
    End With
    
    'オブジェクト解放
    Set INPUT_RANGE = Nothing
    Set CHECK_RANGE = Nothing

End Sub

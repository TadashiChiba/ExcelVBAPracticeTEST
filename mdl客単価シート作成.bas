Attribute VB_Name = "mdl客単価シート作成"
Option Explicit
' ============================================================================
' mdl客単価シート作成
'
' 客単価シートを新規作成し、メインシートの値を元に客単価を計算する場所
' ============================================================================

Private Const PDateRange  As String = "A1"
Private Const PSalesRange As String = "B1"
Private Const PCustomerRange As String = "C1"
Private Const PAverageRange As String = "D1"
Private Const PFormatRange As String = "A:D"
Private Const PDate As String = "日付"
Private Const PSales As String = "売上"
Private Const PCustomer As String = "客数"
Private Const PAverage As String = "客単価"

' ----------------------------------------------------------------------------
' ◆ 客単価シートを作成する
'
' 客単価シート作成リクエストから呼び出される。
' ----------------------------------------------------------------------------
Public Function 客単価シートを作成する(wb As Workbook, ByRef メッセージDe As String) As Boolean
    
    客単価シートを作成する = me前処理を行う(wb)
    If Not 客単価シートを作成する Then
        メッセージDe = "客単価シート作成に失敗しました。"
        Exit Function
    End If

    With wb
        me客単価シートを作成する .Worksheets(Gシート名客単価), .Worksheets(Gシート名メイン)
        me後処理を行う .Worksheets(Gシート名客単価)
    End With

End Function

' ----------------------------------------------------------------------------
' me前処理を行う
'
' 初期化処理(客単価シート削除、新規客単価シート追加,表側を整える)を行う。
' ----------------------------------------------------------------------------
Private Function me前処理を行う(wb As Workbook) As Boolean
    On Error Resume Next
    With wb
        .Worksheets(Gシート名客単価).Delete
        With .Worksheets.Add(after:=Worksheets(Worksheets.Count))
            .Name = Gシート名客単価
            .Range(PDateRange).Value = PDate
            .Range(PSalesRange).Value = PSales
            .Range(PCustomerRange).Value = PCustomer
            .Range(PAverageRange).Value = PAverage
        End With
    End With
    me前処理を行う = Err.Number = 0
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' me客単価シートを作成する
'
' 元となる値をメインシートから取得し設定後する。客単価シート上で客単価を計算する。
' ----------------------------------------------------------------------------
Private Sub me客単価シートを作成する(ws客単価 As Worksheet, wsメイン As Worksheet)

    me客単価計算元値を設定する ws客単価, wsメイン
    me客単価を計算する ws客単価

End Sub

' ----------------------------------------------------------------------------
' me客単価計算元値を設定する
'
' メインシートから客単価シートに計算元データをコピーする。
' ----------------------------------------------------------------------------
Private Sub me客単価計算元値を設定する(ws客単価 As Worksheet, wsメイン As Worksheet)
    Dim i As Long
    Dim j As Long
    Dim 最大行数 As Long

    最大行数 = wsメイン.Cells(Rows.Count, 1).End(xlUp).Row

    With ws客単価
        For i = 2 To 最大行数
            For j = 1 To GC客単価 - 1
                .Cells(i, j).Value = wsメイン.Cells(i, j).Value
            Next
        Next
    End With
End Sub

' ----------------------------------------------------------------------------
' me客単価を計算する
'
' 客単価を計算する
' ----------------------------------------------------------------------------
Private Sub me客単価を計算する(ws客単価 As Worksheet)
    Dim i As Long
    Dim 最大行数 As Long
    
    最大行数 = ws客単価.Cells(Rows.Count, 1).End(xlUp).Row
    
    With ws客単価
        For i = 2 To 最大行数
            me一つの客単価を計算する i
        Next
    End With
End Sub

' ----------------------------------------------------------------------------
' me一つの客単価を計算する
'

' 客単価 = 売上 ÷ 客数。
' ----------------------------------------------------------------------------
Private Sub me一つの客単価を計算する(対象行数)
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets(2)
    With ws
        .Cells(対象行数, GC客単価).Value = .Cells(対象行数, GC売上) \ .Cells(対象行数, GC客数)
    End With
End Sub

' ----------------------------------------------------------------------------
' me後処理を行う
'
' 客単価シートの成形を行う
' ----------------------------------------------------------------------------
Private Sub me後処理を行う(ws As Worksheet)
    ws.Range(PFormatRange).EntireColumn.AutoFit
    With ws.Range(ws.Cells(1, 1), ws.Cells(ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1, GC客単価))
        .CurrentRegion.Borders.LineStyle = xlContinuous
    End With
End Sub

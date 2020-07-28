Attribute VB_Name = "mdl客単価シート作成"
Option Explicit
' ============================================================================
' mdl客単価シート作成
'
' 客単価シートを新規作成し、メインシートの値を元に客単価を計算する場所
' ============================================================================

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
' ◆ me前処理を行う
'
' 初期化処理(客単価シート削除、新規客単価シート追加)を行う。
' ----------------------------------------------------------------------------
Private Function me前処理を行う(wb As Workbook) As Boolean

    me客単価シートを削除する wb
    
    me前処理を行う = me客単価シートを追加する(wb)
    If me前処理を行う Then
        me前処理を行う = True
    End If
    
    me客単価シートに罫線を引く wb
End Function

' ----------------------------------------------------------------------------
' ◆ me客単価シートを削除する
'
' 客単価シートが既に存在していたら削除する
' ----------------------------------------------------------------------------
Private Function me客単価シートを削除する(wb As Workbook)
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = wb.Worksheets(2)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        ws.Delete
    End If
End Function

' ----------------------------------------------------------------------------
' me客単価シートを追加する
'
' 客単価シートを追加する。
' ----------------------------------------------------------------------------
Private Function me客単価シートを追加する(wb As Workbook) As Boolean
    Dim ws As Worksheet

    wb.Worksheets.Add after:=Worksheets(Worksheets.Count)
    
    On Error Resume Next
    Set ws = wb.Worksheets(2)
    
    If Not ws Is Nothing Then
        me客単価シートを追加する = True
    End If
    On Error GoTo 0

    With ws
        .Name = Gシート名客単価
        .Range("A1").Value = "日付"
        .Range("B1").Value = "売上  "
        .Range("C1").Value = "客数"
        .Range("D1").Value = "客単価"
    End With
    
    Set ws = Nothing
End Function

' ----------------------------------------------------------------------------


' ----------------------------------------------------------------------------
Private Sub me客単価シートに罫線を引く(wb As Workbook)
    Dim ws As Worksheet
    Dim 最大行数 As Long
    
    Set ws = wb.Worksheets(2)
    
    最大行数 = wb.Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row

    ws.Range(Cells(1, 1), Cells(最大行数, GC客単価)).Borders.LineStyle = True

End Sub
' ----------------------------------------------------------------------------
' ◆ me客単価シートを作成する
'
' 客単価を計算する。
' ----------------------------------------------------------------------------
Private Sub me客単価シートを作成する(ws As Worksheet, ws2 As Worksheet)
    Dim i As Long
    Dim j As Long
    Dim 最大行数 As Long

    最大行数 = ws2.Cells(Rows.Count, 1).End(xlUp).Row
    
    With ws
        For i = 2 To 最大行数
            For j = 1 To GC客単価 - 1
                .Cells(i, j).Value = ws2.Cells(i, j).Value
            Next
        Next
    End With
    
    With ws
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
   ws.Range("A:D").EntireColumn.AutoFit
End Sub

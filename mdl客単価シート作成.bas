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
    メインシート以外を削除する
    客単価シートを作成する = 客単価を計算する(メッセージDe)
    If Not 客単価シートを作成する Then
        メッセージDe = "メインシートの客数列に空白セルがあるため全ての日付の客単価を計算出来ませんでした。"
    End If
End Function

' ----------------------------------------------------------------------------
' ◆ メインシート以外を削除する
'
' メインシート以外を消去する。
' ----------------------------------------------------------------------------
Private Sub メインシート以外を削除する()
    Dim ws As Worksheet

    For Each ws In Worksheets
        If ws.Name <> Gシート名メイン Then
            ws.Delete
        End If
    Next
End Sub

' ----------------------------------------------------------------------------
' ◆ 客単価を計算する
'
' 客単価を計算する。
' ----------------------------------------------------------------------------
Private Function 客単価を計算する(ByRef メッセージDe As String) As Boolean
    Dim i As Long
    Dim j As Long
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim 最大行数 As Long
    Dim エラー発生行数De As Long
    
    Set ws = ThisWorkbook.Worksheets(1)
    
    最大行数 = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    me客単価シートを追加する (メッセージDe)

    Set ws2 = ThisWorkbook.Worksheets(2)
    
    With ws2
        For i = 2 To 最大行数
            For j = 1 To GC客単価 - 1
                .Cells(i, j).Value = ws.Cells(i, j).Value
                .Cells(i, j).EntireColumn.AutoFit
            Next
        Next
    End With
    
    Worksheets(Gシート名客単価).Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous

    客単価を計算する = me一つの客単価を計算する(ws2, 最大行数, メッセージDe)
   
    Set ws = Nothing
End Function

' ----------------------------------------------------------------------------
' me客単価シートを追加する
'
' 客単価シートを追加する。
' ----------------------------------------------------------------------------
Private Function me客単価シートを追加する(ByRef メッセージDe As String) As Boolean
    Dim ws As Worksheet
    
    Worksheets.Add after:=Worksheets(Worksheets.Count)
    Set ws = ThisWorkbook.Worksheets(2)
    
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
' me一つの客単価を計算する
'
' 客単価 = 売上 ÷ 客数。メインシートの客数列に空白セルがあり0での除算が発生してもその行は無視してエラーメッセージを表示する。
' ----------------------------------------------------------------------------
Private Function me一つの客単価を計算する(ws As Worksheet, 最大行数 As Long, ByRef メッセージDe As String) As Boolean
    Dim i As Long
    Dim j As Long
    
    me一つの客単価を計算する = True
    
    On Error Resume Next
    With ws
        For i = 2 To 最大行数
            .Cells(i, GC客単価).Value = .Cells(i, GC売上) \ .Cells(i, GC客数)
            If Err.Number <> 0 Then
                me一つの客単価を計算する = False
            End If
        Next
    End With
    
    On Error GoTo 0
End Function

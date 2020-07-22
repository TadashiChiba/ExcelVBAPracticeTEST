Attribute VB_Name = "mdl客単価シート作成"
Option Explicit
' ============================================================================
' mdl客単価シート作成
'
' 客単価シートを新規作成し、メインシートの値を元に客単価を計算する場所
' ============================================================================

' ----------------------------------------------------------------------------
' ◆ メインシート以外を削除する
'
' メインシート以外を消去する。
' ----------------------------------------------------------------------------
Public Sub メインシート以外を削除する()
    Dim ws As Worksheet

    Application.DisplayAlerts = False
    For Each ws In Worksheets
        If ws.Name <> Gシート名メイン Then
            ws.Delete
        End If
    Next
    Application.DisplayAlerts = True
End Sub

' ----------------------------------------------------------------------------
' ◆ 客単価を計算する
'
' 客単価を計算する。
' ----------------------------------------------------------------------------
Public Sub 客単価を計算する()
    Dim i As Long
    Dim j As Long
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim 最大行数 As Long
    
    Set ws = ThisWorkbook.Worksheets(1)
    
    最大行数 = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    me客単価シートを追加する

    Set ws2 = ThisWorkbook.Worksheets(2)
    
    With ws2
        For i = 2 To 最大行数
            For j = 1 To GC客単価
                .Cells(i, j).Value = ws.Cells(i, j).Value
                .Cells(i, j).EntireColumn.AutoFit
            Next
        Next
    End With
    
    Worksheets(Gシート名客単価).Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous

    me一つの客単価を計算する ws2, 最大行数
    
    Set ws = Nothing
End Sub

' ----------------------------------------------------------------------------
' me客単価シートを追加する
'
' 客単価シートを追加する。
' ----------------------------------------------------------------------------
Private Sub me客単価シートを追加する()
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
End Sub

' ----------------------------------------------------------------------------
' me一つの客単価を計算する
'
' 客単価 = 売上 ÷ 客数。
' ----------------------------------------------------------------------------
Private Sub me一つの客単価を計算する(ws As Worksheet, 最大行数 As Long)
    Dim i As Long
    Dim j As Long
    
    With ws
        For i = 2 To 最大行数
            .Cells(i, GC客単価).Value = .Cells(i, GC売上) \ .Cells(i, GC客数)
        Next
    End With
End Sub

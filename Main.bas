Attribute VB_Name = "Main"
Option Explicit

Sub startButton_Click()
    Call deleteSeet
    Call getAverageMain
End Sub

'客単価を計算する。
Sub getAverageMain()
    Dim i, j As Integer
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim maxRow As Integer
    
    Set ws = Worksheets(WS_メイン)
    
    maxRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Call setResultSheet

    Set ws2 = Worksheets(WS_客単価)
    
    With ws2
        For i = 2 To maxRow
            For j = 1 To MAXCOLUMN
                .Cells(i, j).Value = ws.Cells(i, j).Value
                .Cells(i, j).EntireColumn.AutoFit
            Next
        Next
    End With
    
    Worksheets(WS_客単価).Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous

    Call getAverage(ws2, maxRow)
End Sub

'客単価 = 売上 ÷ 客数。
Sub getAverage(ws As Worksheet, ByVal maxRow As Integer)
    Dim i As Integer
    Dim j As Integer
    
    With ws
        For i = 2 To maxRow
            .Cells(i, MAXCOLUMN).Value = .Cells(i, SALESCOLUMN) \ .Cells(i, CUSTOMERCOLUMN)
        Next
    End With
End Sub

'メインシート以外を消去。
Sub deleteSeet()
    Dim ws As Worksheet

    Application.DisplayAlerts = False
    For Each ws In Worksheets
        If ws.Name <> WS_メイン Then
            ws.Delete
        End If
    Next
    Application.DisplayAlerts = True
End Sub

'客単価シートを追加。
Sub setResultSheet()
    Dim ws As Worksheet
    
    Set ws = Worksheets.Add(After:=Sheets(Worksheets.Count))

    With ws
        .Name = WS_客単価
        .Range("A1").Value = "日付"
        .Range("B1").Value = "売上  "
        .Range("C1").Value = "客数"
        .Range("D1").Value = "客単価"
    End With
End Sub

Attribute VB_Name = "mdlMain"
Option Explicit
' ============================================================================
' mdlMain
'
' 客単価シートを新規作成し、メインシートの値を元に客単価を計算する。
' ============================================================================

' ----------------------------------------------------------------------------
' ◆ StartButtonClick
'
' メインシート start ボタンクリックで呼び出し。
' ----------------------------------------------------------------------------
Public Sub StartButtonClick()
    meDeleteSeet
    meGetAverageMain
End Sub

' ----------------------------------------------------------------------------
' meGetAverageMain
'
' 客単価を計算する。
' ----------------------------------------------------------------------------
Private Sub meGetAverageMain()
    Dim i As Long
    Dim j As Long
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim maxRow As Long
    
    Set ws = ThisWorkbook.Worksheets(1)
    
    maxRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    meSetResultSheet

    Set ws2 = ThisWorkbook.Worksheets(2)
    
    With ws2
        For i = 2 To maxRow
            For j = 1 To GAverageColumn
                .Cells(i, j).Value = ws.Cells(i, j).Value
                .Cells(i, j).EntireColumn.AutoFit
            Next
        Next
    End With
    
    Worksheets(GWsAverage).Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous

    meGetAverage ws2, maxRow
    
    Set ws = Nothing
End Sub

' ----------------------------------------------------------------------------
' meGetAverage
'
' 客単価 = 売上 ÷ 客数。
' ----------------------------------------------------------------------------
Private Sub meGetAverage(ws As Worksheet, maxRow As Long)
    Dim i As Long
    Dim j As Long
    
    With ws
        For i = 2 To maxRow
            .Cells(i, GAverageColumn).Value = .Cells(i, GSalseColimn) \ .Cells(i, GCustomerColumn)
        Next
    End With
End Sub

' ----------------------------------------------------------------------------
' meDeleteSeet
'
' メインシート以外を消去する。
' ----------------------------------------------------------------------------
Private Sub meDeleteSeet()
    Dim ws As Worksheet

    Application.DisplayAlerts = False
    For Each ws In Worksheets
        If ws.Name <> GWsMain Then
            ws.Delete
        End If
    Next
    Application.DisplayAlerts = True
End Sub

' ----------------------------------------------------------------------------
' meSetResultSheet
'
' 客単価シートを追加する。
' ----------------------------------------------------------------------------
Private Sub meSetResultSheet()
    Dim ws As Worksheet
    
    Worksheets.Add after:=Worksheets(Worksheets.Count)
    
    Set ws = ThisWorkbook.Worksheets(2)

    With ws
        .Name = GWsAverage
        .Range("A1").Value = "日付"
        .Range("B1").Value = "売上  "
        .Range("C1").Value = "客数"
        .Range("D1").Value = "客単価"
    End With
    
    Set ws = Nothing
End Sub

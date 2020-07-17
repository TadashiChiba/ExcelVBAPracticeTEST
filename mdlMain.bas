Attribute VB_Name = "mdlMain"
'============================================================================
'mdlMain
'
'客単価シートを新規作成し、メインシートの値を元に客単価を計算する。
'============================================================================
Option Explicit

Private Sub startButton_Click()
    meDeleteSeet
    meGetAverageMain
End Sub

' ----------------------------------------------------------------------------
'客単価を計算する。
' ----------------------------------------------------------------------------
Private Sub meGetAverageMain()
    Dim i As Long
    Dim j As Long
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim maxRow As Long
    
    Set ws = Worksheets(GWsMain)
    
    maxRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    setResultSheet

    Set ws2 = Worksheets(GWsAverage)
    
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
End Sub

' ----------------------------------------------------------------------------
'客単価 = 売上 ÷ 客数。
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
'メインシート以外を消去する。
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
'客単価シートを追加する。
' ----------------------------------------------------------------------------
Private Sub setResultSheet()
    Dim ws As Worksheet
    
    Set ws = Worksheets.Add(After:=Sheets(Worksheets.Count))

    With ws
        .Name = GWsAverage
        .Range("A1").Value = "日付"
        .Range("B1").Value = "売上  "
        .Range("C1").Value = "客数"
        .Range("D1").Value = "客単価"
    End With
End Sub

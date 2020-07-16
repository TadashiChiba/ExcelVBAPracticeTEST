Attribute VB_Name = "Main"
Option Explicit

Sub startButton_Click()
    Call deleteSeet
    Call getAverageMain
End Sub

'�q�P�����v�Z����B
Sub getAverageMain()
    Dim i, j As Integer
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim maxRow As Integer
    
    Set ws = Worksheets(WS_���C��)
    
    maxRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Call setResultSheet

    Set ws2 = Worksheets(WS_�q�P��)
    
    With ws2
        For i = 2 To maxRow
            For j = 1 To MAXCOLUMN
                .Cells(i, j).Value = ws.Cells(i, j).Value
                .Cells(i, j).EntireColumn.AutoFit
            Next
        Next
    End With
    
    Worksheets(WS_�q�P��).Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous

    Call getAverage(ws2, maxRow)
End Sub

'�q�P�� = ���� �� �q���B
Sub getAverage(ws As Worksheet, ByVal maxRow As Integer)
    Dim i As Integer
    Dim j As Integer
    
    With ws
        For i = 2 To maxRow
            .Cells(i, MAXCOLUMN).Value = .Cells(i, SALESCOLUMN) \ .Cells(i, CUSTOMERCOLUMN)
        Next
    End With
End Sub

'���C���V�[�g�ȊO�������B
Sub deleteSeet()
    Dim ws As Worksheet

    Application.DisplayAlerts = False
    For Each ws In Worksheets
        If ws.Name <> WS_���C�� Then
            ws.Delete
        End If
    Next
    Application.DisplayAlerts = True
End Sub

'�q�P���V�[�g��ǉ��B
Sub setResultSheet()
    Dim ws As Worksheet
    
    Set ws = Worksheets.Add(After:=Sheets(Worksheets.Count))

    With ws
        .Name = WS_�q�P��
        .Range("A1").Value = "���t"
        .Range("B1").Value = "����  "
        .Range("C1").Value = "�q��"
        .Range("D1").Value = "�q�P��"
    End With
End Sub

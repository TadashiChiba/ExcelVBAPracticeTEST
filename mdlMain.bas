Attribute VB_Name = "mdlMain"
'============================================================================
'mdlMain
'
'�q�P���V�[�g��V�K�쐬���A���C���V�[�g�̒l�����ɋq�P�����v�Z����B
'============================================================================
Option Explicit

Private Sub startButton_Click()
    meDeleteSeet
    meGetAverageMain
End Sub

' ----------------------------------------------------------------------------
'�q�P�����v�Z����B
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
'�q�P�� = ���� �� �q���B
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
'���C���V�[�g�ȊO����������B
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
'�q�P���V�[�g��ǉ�����B
' ----------------------------------------------------------------------------
Private Sub setResultSheet()
    Dim ws As Worksheet
    
    Set ws = Worksheets.Add(After:=Sheets(Worksheets.Count))

    With ws
        .Name = GWsAverage
        .Range("A1").Value = "���t"
        .Range("B1").Value = "����  "
        .Range("C1").Value = "�q��"
        .Range("D1").Value = "�q�P��"
    End With
End Sub

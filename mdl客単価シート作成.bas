Attribute VB_Name = "mdl�q�P���V�[�g�쐬"
Option Explicit
' ============================================================================
' mdl�q�P���V�[�g�쐬
'
' �q�P���V�[�g��V�K�쐬���A���C���V�[�g�̒l�����ɋq�P�����v�Z����ꏊ
' ============================================================================

' ----------------------------------------------------------------------------
' �� ���C���V�[�g�ȊO���폜����
'
' ���C���V�[�g�ȊO����������B
' ----------------------------------------------------------------------------
Public Sub ���C���V�[�g�ȊO���폜����()
    Dim ws As Worksheet

    Application.DisplayAlerts = False
    For Each ws In Worksheets
        If ws.Name <> G�V�[�g�����C�� Then
            ws.Delete
        End If
    Next
    Application.DisplayAlerts = True
End Sub

' ----------------------------------------------------------------------------
' �� �q�P�����v�Z����
'
' �q�P�����v�Z����B
' ----------------------------------------------------------------------------
Public Sub �q�P�����v�Z����()
    Dim i As Long
    Dim j As Long
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim �ő�s�� As Long
    
    Set ws = ThisWorkbook.Worksheets(1)
    
    �ő�s�� = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    me�q�P���V�[�g��ǉ�����

    Set ws2 = ThisWorkbook.Worksheets(2)
    
    With ws2
        For i = 2 To �ő�s��
            For j = 1 To GC�q�P��
                .Cells(i, j).Value = ws.Cells(i, j).Value
                .Cells(i, j).EntireColumn.AutoFit
            Next
        Next
    End With
    
    Worksheets(G�V�[�g���q�P��).Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous

    me��̋q�P�����v�Z���� ws2, �ő�s��
    
    Set ws = Nothing
End Sub

' ----------------------------------------------------------------------------
' me�q�P���V�[�g��ǉ�����
'
' �q�P���V�[�g��ǉ�����B
' ----------------------------------------------------------------------------
Private Sub me�q�P���V�[�g��ǉ�����()
    Dim ws As Worksheet
    
    Worksheets.Add after:=Worksheets(Worksheets.Count)
    
    Set ws = ThisWorkbook.Worksheets(2)

    With ws
        .Name = G�V�[�g���q�P��
        .Range("A1").Value = "���t"
        .Range("B1").Value = "����  "
        .Range("C1").Value = "�q��"
        .Range("D1").Value = "�q�P��"
    End With
    
    Set ws = Nothing
End Sub

' ----------------------------------------------------------------------------
' me��̋q�P�����v�Z����
'
' �q�P�� = ���� �� �q���B
' ----------------------------------------------------------------------------
Private Sub me��̋q�P�����v�Z����(ws As Worksheet, �ő�s�� As Long)
    Dim i As Long
    Dim j As Long
    
    With ws
        For i = 2 To �ő�s��
            .Cells(i, GC�q�P��).Value = .Cells(i, GC����) \ .Cells(i, GC�q��)
        Next
    End With
End Sub

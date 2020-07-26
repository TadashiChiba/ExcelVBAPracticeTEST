Attribute VB_Name = "mdl�q�P���V�[�g�쐬"
Option Explicit
' ============================================================================
' mdl�q�P���V�[�g�쐬
'
' �q�P���V�[�g��V�K�쐬���A���C���V�[�g�̒l�����ɋq�P�����v�Z����ꏊ
' ============================================================================

' ----------------------------------------------------------------------------
' �� �q�P���V�[�g���쐬����
'
' �q�P���V�[�g�쐬���N�G�X�g����Ăяo�����B
' ----------------------------------------------------------------------------
Public Function �q�P���V�[�g���쐬����(wb As Workbook, ByRef ���b�Z�[�W As String) As Boolean
    
    �q�P���V�[�g���쐬���� = True
    
    ���C���V�[�g�ȊO���폜����
    �q�P���V�[�g���쐬���� = �q�P�����v�Z����(���b�Z�[�W)
End Function

' ----------------------------------------------------------------------------
' �� ���C���V�[�g�ȊO���폜����
'
' ���C���V�[�g�ȊO����������B
' ----------------------------------------------------------------------------
Private Sub ���C���V�[�g�ȊO���폜����()
    Dim ws As Worksheet

    For Each ws In Worksheets
        If ws.Name <> G�V�[�g�����C�� Then
            ws.Delete
        End If
    Next
End Sub

' ----------------------------------------------------------------------------
' �� �q�P�����v�Z����
'
' �q�P�����v�Z����B
' ----------------------------------------------------------------------------
Private Function �q�P�����v�Z����(ByRef ���b�Z�[�W As String) As Boolean
    Dim i As Long
    Dim j As Long
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim �ő�s�� As Long
    
    �q�P�����v�Z���� = True
    
    Set ws = ThisWorkbook.Worksheets(1)
    
    �ő�s�� = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    �q�P�����v�Z���� = me�q�P���V�[�g��ǉ�����(���b�Z�[�W)

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
End Function

' ----------------------------------------------------------------------------
' me�q�P���V�[�g��ǉ�����
'
' �q�P���V�[�g��ǉ�����B
' ----------------------------------------------------------------------------
Private Function me�q�P���V�[�g��ǉ�����(ByRef ���b�Z�[�W As String) As Boolean
    Dim ws As Worksheet
    
    me�q�P���V�[�g��ǉ����� = True

    Worksheets.Add after:=Worksheets(Worksheets.Count)
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(2)
    
    If ws Is Nothing Then
        me�q�P���V�[�g��ǉ����� = False
        ���b�Z�[�W = "�V�[�g" & G�V�[�g���q�P�� & "�̒ǉ��Ɏ��s���܂����B"
    End If
    On Error GoTo 0

    With ws
        .Name = G�V�[�g���q�P��
        .Range("A1").Value = "���t"
        .Range("B1").Value = "����  "
        .Range("C1").Value = "�q��"
        .Range("D1").Value = "�q�P��"
    End With
    
    Set ws = Nothing
End Function

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

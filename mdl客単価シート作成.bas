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
Public Function �q�P���V�[�g���쐬����(wb As Workbook, ByRef ���b�Z�[�WDe As String) As Boolean
    ���C���V�[�g�ȊO���폜����
    �q�P���V�[�g���쐬���� = �q�P�����v�Z����(���b�Z�[�WDe)
    If Not �q�P���V�[�g���쐬���� Then
        ���b�Z�[�WDe = "���C���V�[�g�̋q����ɋ󔒃Z�������邽�ߑS�Ă̓��t�̋q�P�����v�Z�o���܂���ł����B"
    End If
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
Private Function �q�P�����v�Z����(ByRef ���b�Z�[�WDe As String) As Boolean
    Dim i As Long
    Dim j As Long
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim �ő�s�� As Long
    Dim �G���[�����s��De As Long
    
    Set ws = ThisWorkbook.Worksheets(1)
    
    �ő�s�� = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    me�q�P���V�[�g��ǉ����� (���b�Z�[�WDe)

    Set ws2 = ThisWorkbook.Worksheets(2)
    
    With ws2
        For i = 2 To �ő�s��
            For j = 1 To GC�q�P�� - 1
                .Cells(i, j).Value = ws.Cells(i, j).Value
                .Cells(i, j).EntireColumn.AutoFit
            Next
        Next
    End With
    
    Worksheets(G�V�[�g���q�P��).Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous

    �q�P�����v�Z���� = me��̋q�P�����v�Z����(ws2, �ő�s��, ���b�Z�[�WDe)
   
    Set ws = Nothing
End Function

' ----------------------------------------------------------------------------
' me�q�P���V�[�g��ǉ�����
'
' �q�P���V�[�g��ǉ�����B
' ----------------------------------------------------------------------------
Private Function me�q�P���V�[�g��ǉ�����(ByRef ���b�Z�[�WDe As String) As Boolean
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
End Function

' ----------------------------------------------------------------------------
' me��̋q�P�����v�Z����
'
' �q�P�� = ���� �� �q���B���C���V�[�g�̋q����ɋ󔒃Z��������0�ł̏��Z���������Ă����̍s�͖������ăG���[���b�Z�[�W��\������B
' ----------------------------------------------------------------------------
Private Function me��̋q�P�����v�Z����(ws As Worksheet, �ő�s�� As Long, ByRef ���b�Z�[�WDe As String) As Boolean
    Dim i As Long
    Dim j As Long
    
    me��̋q�P�����v�Z���� = True
    
    On Error Resume Next
    With ws
        For i = 2 To �ő�s��
            .Cells(i, GC�q�P��).Value = .Cells(i, GC����) \ .Cells(i, GC�q��)
            If Err.Number <> 0 Then
                me��̋q�P�����v�Z���� = False
            End If
        Next
    End With
    
    On Error GoTo 0
End Function

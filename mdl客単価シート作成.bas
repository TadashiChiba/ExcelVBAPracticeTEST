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
    
    �q�P���V�[�g���쐬���� = me�O�������s��(wb)
    If Not �q�P���V�[�g���쐬���� Then
        ���b�Z�[�WDe = "�q�P���V�[�g�쐬�Ɏ��s���܂����B"
        Exit Function
    End If

    With wb
        me�q�P���V�[�g���쐬���� .Worksheets(G�V�[�g���q�P��), .Worksheets(G�V�[�g�����C��)
        me�㏈�����s�� .Worksheets(G�V�[�g���q�P��)
    End With

End Function

' ----------------------------------------------------------------------------
' �� me�O�������s��
'
' ����������(�q�P���V�[�g�폜�A�V�K�q�P���V�[�g�ǉ�)���s���B
' ----------------------------------------------------------------------------
Private Function me�O�������s��(wb As Workbook) As Boolean

    me�q�P���V�[�g���폜���� wb
    
    me�O�������s�� = me�q�P���V�[�g��ǉ�����(wb)
    If me�O�������s�� Then
        me�O�������s�� = True
    End If
    
    me�q�P���V�[�g�Ɍr�������� wb
End Function

' ----------------------------------------------------------------------------
' �� me�q�P���V�[�g���폜����
'
' �q�P���V�[�g�����ɑ��݂��Ă�����폜����
' ----------------------------------------------------------------------------
Private Function me�q�P���V�[�g���폜����(wb As Workbook)
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = wb.Worksheets(2)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        ws.Delete
    End If
End Function

' ----------------------------------------------------------------------------
' me�q�P���V�[�g��ǉ�����
'
' �q�P���V�[�g��ǉ�����B
' ----------------------------------------------------------------------------
Private Function me�q�P���V�[�g��ǉ�����(wb As Workbook) As Boolean
    Dim ws As Worksheet

    wb.Worksheets.Add after:=Worksheets(Worksheets.Count)
    
    On Error Resume Next
    Set ws = wb.Worksheets(2)
    
    If Not ws Is Nothing Then
        me�q�P���V�[�g��ǉ����� = True
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


' ----------------------------------------------------------------------------
Private Sub me�q�P���V�[�g�Ɍr��������(wb As Workbook)
    Dim ws As Worksheet
    Dim �ő�s�� As Long
    
    Set ws = wb.Worksheets(2)
    
    �ő�s�� = wb.Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row

    ws.Range(Cells(1, 1), Cells(�ő�s��, GC�q�P��)).Borders.LineStyle = True

End Sub
' ----------------------------------------------------------------------------
' �� me�q�P���V�[�g���쐬����
'
' �q�P�����v�Z����B
' ----------------------------------------------------------------------------
Private Sub me�q�P���V�[�g���쐬����(ws As Worksheet, ws2 As Worksheet)
    Dim i As Long
    Dim j As Long
    Dim �ő�s�� As Long

    �ő�s�� = ws2.Cells(Rows.Count, 1).End(xlUp).Row
    
    With ws
        For i = 2 To �ő�s��
            For j = 1 To GC�q�P�� - 1
                .Cells(i, j).Value = ws2.Cells(i, j).Value
            Next
        Next
    End With
    
    With ws
        For i = 2 To �ő�s��
            me��̋q�P�����v�Z���� i
        Next
    End With
End Sub

' ----------------------------------------------------------------------------
' me��̋q�P�����v�Z����
'

' �q�P�� = ���� �� �q���B
' ----------------------------------------------------------------------------
Private Sub me��̋q�P�����v�Z����(�Ώۍs��)
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets(2)
    With ws
        .Cells(�Ώۍs��, GC�q�P��).Value = .Cells(�Ώۍs��, GC����) \ .Cells(�Ώۍs��, GC�q��)
    End With
End Sub

' ----------------------------------------------------------------------------
' me�㏈�����s��
'
' �q�P���V�[�g�̐��`���s��
' ----------------------------------------------------------------------------
Private Sub me�㏈�����s��(ws As Worksheet)
   ws.Range("A:D").EntireColumn.AutoFit
End Sub

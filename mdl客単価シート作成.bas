Attribute VB_Name = "mdl�q�P���V�[�g�쐬"
Option Explicit
' ============================================================================
' mdl�q�P���V�[�g�쐬
'
' �q�P���V�[�g��V�K�쐬���A���C���V�[�g�̒l�����ɋq�P�����v�Z����ꏊ
' ============================================================================

Private Const PDateRange  As String = "A1"
Private Const PSalesRange As String = "B1"
Private Const PCustomerRange As String = "C1"
Private Const PAverageRange As String = "D1"
Private Const PFormatRange As String = "A:D"
Private Const PDate As String = "���t"
Private Const PSales As String = "����"
Private Const PCustomer As String = "�q��"
Private Const PAverage As String = "�q�P��"

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
' me�O�������s��
'
' ����������(�q�P���V�[�g�폜�A�V�K�q�P���V�[�g�ǉ�,�\���𐮂���)���s���B
' ----------------------------------------------------------------------------
Private Function me�O�������s��(wb As Workbook) As Boolean
    On Error Resume Next
    With wb
        .Worksheets(G�V�[�g���q�P��).Delete
        With .Worksheets.Add(after:=Worksheets(Worksheets.Count))
            .Name = G�V�[�g���q�P��
            .Range(PDateRange).Value = PDate
            .Range(PSalesRange).Value = PSales
            .Range(PCustomerRange).Value = PCustomer
            .Range(PAverageRange).Value = PAverage
        End With
    End With
    me�O�������s�� = Err.Number = 0
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' me�q�P���V�[�g���쐬����
'
' ���ƂȂ�l�����C���V�[�g����擾���ݒ�シ��B�q�P���V�[�g��ŋq�P�����v�Z����B
' ----------------------------------------------------------------------------
Private Sub me�q�P���V�[�g���쐬����(ws�q�P�� As Worksheet, ws���C�� As Worksheet)

    me�q�P���v�Z���l��ݒ肷�� ws�q�P��, ws���C��
    me�q�P�����v�Z���� ws�q�P��

End Sub

' ----------------------------------------------------------------------------
' me�q�P���v�Z���l��ݒ肷��
'
' ���C���V�[�g����q�P���V�[�g�Ɍv�Z���f�[�^���R�s�[����B
' ----------------------------------------------------------------------------
Private Sub me�q�P���v�Z���l��ݒ肷��(ws�q�P�� As Worksheet, ws���C�� As Worksheet)
    Dim i As Long
    Dim j As Long
    Dim �ő�s�� As Long

    �ő�s�� = ws���C��.Cells(Rows.Count, 1).End(xlUp).Row

    With ws�q�P��
        For i = 2 To �ő�s��
            For j = 1 To GC�q�P�� - 1
                .Cells(i, j).Value = ws���C��.Cells(i, j).Value
            Next
        Next
    End With
End Sub

' ----------------------------------------------------------------------------
' me�q�P�����v�Z����
'
' �q�P�����v�Z����
' ----------------------------------------------------------------------------
Private Sub me�q�P�����v�Z����(ws�q�P�� As Worksheet)
    Dim i As Long
    Dim �ő�s�� As Long
    
    �ő�s�� = ws�q�P��.Cells(Rows.Count, 1).End(xlUp).Row
    
    With ws�q�P��
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
    ws.Range(PFormatRange).EntireColumn.AutoFit
    With ws.Range(ws.Cells(1, 1), ws.Cells(ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1, GC�q�P��))
        .CurrentRegion.Borders.LineStyle = xlContinuous
    End With
End Sub

Attribute VB_Name = "mdlMenu"
Option Explicit
' ============================================================================
' mdlMenu
'
' ���[�U����̃C�x���g���󂯕t����ꏊ�B
' ============================================================================

' ----------------------------------------------------------------------------
' �� bt�q�P���V�[�g���쐬����
'
' ���C���V�[�g start �{�^���N���b�N�C�x���g���󂯕t����B
' ----------------------------------------------------------------------------
Public Sub bt�q�P���V�[�g�쐬()
    Dim Succeed As Boolean
    Dim Message As String
    
    With Application
        .Cursor = xlWait
        .DisplayAlerts = False
        .ScreenUpdating = False
        Succeed = �q�P���V�[�g���쐬����(ThisWorkbook, Message)
        .Cursor = xlDefault
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
    If Not Succeed Then
        MsgBox Message, vbExclamation
'        MsgBox Message, vbExclamation, sySystemTitle
    End If
    
End Sub

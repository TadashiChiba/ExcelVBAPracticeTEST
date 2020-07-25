Attribute VB_Name = "mdlMenu"
Option Explicit
' ============================================================================
' mdlMenu
'
' ユーザからのイベントを受け付ける場所。
' ============================================================================

' ----------------------------------------------------------------------------
' ◆ bt客単価シートを作成する
'
' メインシート start ボタンクリックイベントを受け付ける。
' ----------------------------------------------------------------------------
Public Sub bt客単価シート作成()
    Dim Succeed As Boolean
    Dim Message As String
    
    With Application
        .Cursor = xlWait
        .DisplayAlerts = False
        .ScreenUpdating = False
        Succeed = 客単価シートを作成する(ThisWorkbook, Message)
        .Cursor = xlDefault
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
    If Not Succeed Then
        MsgBox Message, vbExclamation
'        MsgBox Message, vbExclamation, sySystemTitle
    End If
    
End Sub

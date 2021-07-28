VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------------
' Sub ShowError
'
' �G���[���b�Z�[�W��\��
'
' In     : [vstrMessage] - ���b�Z�[�W���e
' Out    :
' Return :
'-----------------------------------------------------------
Public Sub ShowError(ByVal vstrMessage As String)

    Call MsgBox(vstrMessage, vbOKOnly + vbCritical, "�G���[")

End Sub

'-----------------------------------------------------------
' Sub ShowMessage
'
' �x�����b�Z�[�W��\��
'
' In     : [vstrMessage] - ���b�Z�[�W���e
' Out    :
' Return :
'-----------------------------------------------------------
Public Sub ShowMessage(ByVal vstrMessage As String)

    Call MsgBox(vstrMessage, vbOKOnly + vbExclamation, "�x��")

End Sub

'-----------------------------------------------------------
' Sub ShowInformation
'
' ��񃁃b�Z�[�W��\��
'
' In     : [vstrInformation] - ���b�Z�[�W���e
' Out    :
' Return :
'-----------------------------------------------------------
Public Sub ShowInformation(ByVal vstrInformation As String)

    Call MsgBox(vstrInformation, vbOKOnly + vbInformation, "���")

End Sub

'-----------------------------------------------------------
' Function ShowQuestion
'
' ���⃁�b�Z�[�W��\��
'
' In     : [vstrQuestion] - ���b�Z�[�W���e
' Out    :
' Return : MsgBox�߂�l
'-----------------------------------------------------------
Public Function ShowQuestion(ByVal vstrQuestion As String) As Integer

    ShowQuestion = MsgBox(vstrQuestion, vbYesNo + vbQuestion + vbDefaultButton2, "����")

End Function

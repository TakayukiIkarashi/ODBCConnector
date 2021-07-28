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
' エラーメッセージを表示
'
' In     : [vstrMessage] - メッセージ内容
' Out    :
' Return :
'-----------------------------------------------------------
Public Sub ShowError(ByVal vstrMessage As String)

    Call MsgBox(vstrMessage, vbOKOnly + vbCritical, "エラー")

End Sub

'-----------------------------------------------------------
' Sub ShowMessage
'
' 警告メッセージを表示
'
' In     : [vstrMessage] - メッセージ内容
' Out    :
' Return :
'-----------------------------------------------------------
Public Sub ShowMessage(ByVal vstrMessage As String)

    Call MsgBox(vstrMessage, vbOKOnly + vbExclamation, "警告")

End Sub

'-----------------------------------------------------------
' Sub ShowInformation
'
' 情報メッセージを表示
'
' In     : [vstrInformation] - メッセージ内容
' Out    :
' Return :
'-----------------------------------------------------------
Public Sub ShowInformation(ByVal vstrInformation As String)

    Call MsgBox(vstrInformation, vbOKOnly + vbInformation, "情報")

End Sub

'-----------------------------------------------------------
' Function ShowQuestion
'
' 質問メッセージを表示
'
' In     : [vstrQuestion] - メッセージ内容
' Out    :
' Return : MsgBox戻り値
'-----------------------------------------------------------
Public Function ShowQuestion(ByVal vstrQuestion As String) As Integer

    ShowQuestion = MsgBox(vstrQuestion, vbYesNo + vbQuestion + vbDefaultButton2, "質問")

End Function


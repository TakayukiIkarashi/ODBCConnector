VERSION 5.00
Begin VB.Form frmTimeout 
   BorderStyle     =   4  '固定ﾂｰﾙ ｳｨﾝﾄﾞｳ
   Caption         =   "クエリ タイムアウト"
   ClientHeight    =   1920
   ClientLeft      =   2850
   ClientTop       =   1785
   ClientWidth     =   4905
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Meiryo UI"
      Size            =   9.75
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox txtTimeout 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Left            =   3420
      MaxLength       =   8
      TabIndex        =   1
      Text            =   "12345678"
      ToolTipText     =   "秒単位で指定"
      Top             =   840
      Width           =   1215
   End
   Begin VB.Frame fraBtn 
      BorderStyle     =   0  'なし
      Caption         =   "Frame5"
      Height          =   555
      Left            =   1620
      TabIndex        =   2
      Top             =   1320
      Width           =   3255
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "キャンセル"
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "frmTimeout.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblMsg 
      Caption         =   "クエリ タイムアウト時間を入力"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "frmTimeout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------
' Event Form_Load()
'------------------------------------------------------------
Public Sub Form_Load()
    Dim intPos                  As Integer

    With txtTimeout
        .Text = CStr(pobjCommand.CommandTimeout)
        intPos = Len(.Text)
        .SelStart = intPos
        .SelLength = 0
    End With

End Sub

'------------------------------------------------------------
' Event cmdOK_Click()
'------------------------------------------------------------
Private Sub cmdOK_Click()

    pobjCommand.CommandTimeout = CLng(txtTimeout.Text)

    Unload Me

End Sub

'------------------------------------------------------------
' Event cmdCancel_Click()
'------------------------------------------------------------
Private Sub cmdCancel_Click()

    Unload Me

End Sub

'------------------------------------------------------------
' Event txtTimeout_Change()
'------------------------------------------------------------
Private Sub txtTimeout_Change()
    Dim strWrk                  As String
    Dim intPos                  As Integer

    With txtTimeout
        If .Text = vbNullString Then
            .Text = "0"
        End If
        If (Len(.Text) > 1) And (Left(.Text, 1) = "0") Then
            .Text = Right(.Text, Len(.Text) - 1)
        End If
        intPos = Len(.Text)
        .SelStart = intPos
        .SelLength = 0
    End With

End Sub

'------------------------------------------------------------
' Event txtTimeout_KeyDown()
'------------------------------------------------------------
Private Sub txtTimeout_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyDelete
            Call SendKeys("{BS}")

        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
            KeyCode = vbKeyCancel
    End Select

End Sub

'------------------------------------------------------------
' Event txtTimeout_KeyPress()
'------------------------------------------------------------
Private Sub txtTimeout_KeyPress(KeyAscii As Integer)
    Dim lngWrk                  As Long
    Dim strWrk                  As String

    Select Case KeyAscii
        '"BS"
        Case 8

        '"0","1","2","3","4","5","6","7","8","9"
        Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57

        Case Else
            KeyAscii = vbKeyCancel
    End Select

End Sub

'------------------------------------------------------------
' Event txtTimeout_MouseUp()
'------------------------------------------------------------
Private Sub txtTimeout_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intPos                  As Integer

    With txtTimeout
        intPos = Len(.Text)
        .SelStart = intPos
        .SelLength = 0
    End With

End Sub


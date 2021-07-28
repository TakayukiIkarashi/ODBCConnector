VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  '�Œ�°� ����޳
   Caption         =   "�o�[�W�������"
   ClientHeight    =   2055
   ClientLeft      =   2340
   ClientTop       =   1890
   ClientWidth     =   5235
   ClipControls    =   0   'False
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
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1418.398
   ScaleMode       =   0  'հ�ް
   ScaleWidth      =   4915.936
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":000C
      ScaleHeight     =   337.12
      ScaleMode       =   0  'հ�ް
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3600
      TabIndex        =   0
      Top             =   1500
      Width           =   1380
   End
   Begin VB.Label lblAuthor 
      Alignment       =   1  '�E����
      Caption         =   "https://ikachi.org"
      BeginProperty Font 
         Name            =   "Meiryo UI"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1050
      TabIndex        =   4
      Top             =   1080
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "���ع���� ����"
      BeginProperty Font 
         Name            =   "Meiryo UI"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   2
      Top             =   240
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  '��������
      Caption         =   "�ް�ޮ�"
      Height          =   255
      Left            =   1050
      TabIndex        =   3
      Top             =   780
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------
' Event Form_Load()
'------------------------------------------------------------
Private Sub Form_Load()

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title

End Sub

'------------------------------------------------------------
' Event cmdOK_Click()
'------------------------------------------------------------
Private Sub cmdOK_Click()

    Unload Me

End Sub

'------------------------------------------------------------
' Event cmdOK_Click()
'------------------------------------------------------------
Private Sub lblAuthor_Click()

    Call ShellExecute(Me.hwnd, "Open", "https://ikachi.org/", vbNullString, App.Path, 1)

End Sub

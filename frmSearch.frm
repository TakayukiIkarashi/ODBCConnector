VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   4  'å≈íË¬∞Ÿ ≥®›ƒﬁ≥
   Caption         =   "åüçı"
   ClientHeight    =   1740
   ClientLeft      =   165
   ClientTop       =   285
   ClientWidth     =   6765
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
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'µ∞≈∞ Ã´∞—ÇÃíÜâõ
   Begin VB.Frame fraBtn 
      BorderStyle     =   0  'Ç»Çµ
      Caption         =   "Frame1"
      Height          =   855
      Left            =   5220
      TabIndex        =   3
      Top             =   840
      Width           =   1515
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ÉLÉÉÉìÉZÉã"
         Height          =   375
         Left            =   30
         TabIndex        =   4
         Top             =   450
         Width           =   1455
      End
   End
   Begin VB.CheckBox chkUpperConv 
      Appearance      =   0  'Ã◊Øƒ
      BackColor       =   &H80000005&
      Caption         =   "ëÂï∂éöè¨ï∂éöÇãÊï (&C)"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   300
      TabIndex        =   1
      Top             =   1350
      Width           =   4485
   End
   Begin VB.TextBox txtWord 
      Appearance      =   0  'Ã◊Øƒ
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Text            =   "1234567890123456789012345678901234567890"
      Top             =   870
      Width           =   3675
   End
   Begin VB.Label lblMsg 
      Caption         =   "ï∂éöóÒÇåüçı"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   360
      Width           =   5895
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "frmSearch.frx":000C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblSearch 
      Caption         =   "åüçı(&N)ÅF"
      Height          =   285
      Left            =   300
      TabIndex        =   2
      Top             =   900
      Width           =   1095
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------
' Event Form_Load()
'------------------------------------------------------------
Private Sub Form_Load()
    Dim strRet                  As String

    On Error Resume Next

    txtWord.Text = pstrSearchWord

    strRet = GetSetting(App.Title, SEC_SEARCH, KEY_UPPER, CStr(vbUnchecked))

    chkUpperConv.Value = CLng(strRet)

    Err.Clear

End Sub

'------------------------------------------------------------
' Event cmdOK_Click()
'------------------------------------------------------------
Private Sub cmdOK_Click()

    Call SaveSetting(App.Title, SEC_SEARCH, KEY_UPPER, CStr(chkUpperConv.Value))

    pstrSearchWord = txtWord.Text
    pblnUpperConv = (chkUpperConv.Value <> vbChecked)

    Unload Me

End Sub

'------------------------------------------------------------
' Event cmdCancel_Click()
'------------------------------------------------------------
Private Sub cmdCancel_Click()

    Unload Me

End Sub


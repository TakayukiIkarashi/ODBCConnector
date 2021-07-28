VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   4  '固定ﾂｰﾙ ｳｨﾝﾄﾞｳ
   Caption         =   "データ ソース接続"
   ClientHeight    =   3285
   ClientLeft      =   2850
   ClientTop       =   1785
   ClientWidth     =   6735
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
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame fraBtn 
      BorderStyle     =   0  'なし
      Caption         =   "Frame5"
      Height          =   555
      Left            =   3450
      TabIndex        =   8
      Top             =   2700
      Width           =   3255
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "キャンセル"
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame fraDSN 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   6465
      Begin VB.TextBox txtUID 
         Appearance      =   0  'ﾌﾗｯﾄ
         Height          =   375
         Left            =   2370
         TabIndex        =   5
         Text            =   "12345678901234567890123456789012"
         Top             =   900
         Width           =   3675
      End
      Begin VB.TextBox txtPWD 
         Appearance      =   0  'ﾌﾗｯﾄ
         Height          =   375
         IMEMode         =   3  'ｵﾌ固定
         Left            =   2370
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1320
         Width           =   3675
      End
      Begin VB.ComboBox cboDSNList 
         Appearance      =   0  'ﾌﾗｯﾄ
         Height          =   375
         ItemData        =   "frmConnect.frx":000C
         Left            =   2370
         List            =   "frmConnect.frx":000E
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "cboDSNList"
         Top             =   420
         Width           =   3720
      End
      Begin VB.Label lblDSN 
         Alignment       =   1  '右揃え
         Caption         =   "データソース名(&S)："
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   2085
      End
      Begin VB.Label lblUID 
         Alignment       =   1  '右揃え
         Caption         =   "ログイン名(&L)："
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   900
         Width           =   2085
      End
      Begin VB.Label lblPWD 
         Alignment       =   1  '右揃え
         Caption         =   "パスワード(&P)："
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   2085
      End
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "frmConnect.frx":0010
      Stretch         =   -1  'True
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblMsg 
      Caption         =   "データソース接続情報"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   5745
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------
'   API関数定義
'------------------------------
Private Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer
Private Declare Function SQLAllocEnv% Lib "ODBC32.DLL" (env&)


'------------------------------
'   定数定義
'------------------------------
Private Const SQL_SUCCESS       As Long = 0
Private Const SQL_FETCH_NEXT    As Long = 1

Private Const SEC_CONNECT       As String = "CONNECT"
Private Const KEY_DSN           As String = "DSN"
Private Const KEY_UID           As String = "UID"


'------------------------------
'   変数定義
'------------------------------
Private mclsMsg                 As New clsMsg                   'メッセージ出力クラス

'------------------------------------------------------------
' Event Form_Load()
'------------------------------------------------------------
Private Sub Form_Load()

    Call GetDSNsAndDrivers

    cboDSNList.Text = GetSetting(App.Title, SEC_CONNECT, KEY_DSN)

    txtUID.Text = GetSetting(App.Title, SEC_CONNECT, KEY_UID)

End Sub

'------------------------------------------------------------
' Event Form_QueryUnload()
'------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call SaveSetting(App.Title, SEC_CONNECT, KEY_DSN, cboDSNList.Text)
    Call SaveSetting(App.Title, SEC_CONNECT, KEY_UID, txtUID.Text)

End Sub

'------------------------------------------------------------
' Event cmdOK_Click()
'------------------------------------------------------------
Private Sub cmdOK_Click()

    If Not (Connect) Then Exit Sub

    pstrDSN = cboDSNList.Text

    Unload Me

End Sub

'------------------------------------------------------------
' Event cmdCancel_Click()
'------------------------------------------------------------
Private Sub cmdCancel_Click()

    Unload Me

End Sub

'------------------------------------------------------------
' Sub GetDSNsAndDrivers()
'
' データ ソースをコンボボックスにセット
'
' In     : なし
' Out    : なし
' Return : なし
'------------------------------------------------------------
Sub GetDSNsAndDrivers()
    Dim i                       As Integer
    Dim sDSNItem                As String * 1024
    Dim sDRVItem                As String * 1024
    Dim sDSN                    As String
    Dim iDSNLen                 As Integer
    Dim iDRVLen                 As Integer
    Dim lHenv                   As Long

    On Error Resume Next

    If SQLAllocEnv(lHenv) <> -1 Then
        Do Until i <> SQL_SUCCESS
            sDSNItem = Space$(1024)
            sDRVItem = Space$(1024)
            i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
            sDSN = Left$(sDSNItem, iDSNLen)

            If sDSN <> Space(iDSNLen) Then
                cboDSNList.AddItem sDSN
            End If
        Loop
    End If

    Err.Clear

End Sub

'------------------------------------------------------------
' Function Connect()
'
' 接続処理
'
' In     : なし
' Out    : なし
' Return : 接続できたらTrue、そうでなければFalse
'------------------------------------------------------------
Private Function Connect() As Boolean
    Dim sConnect                As String

    On Error GoTo Exception

    Connect = False

    If (pobjConnection.State <> 0) Then pobjConnection.Close

    sConnect = vbNullString
    sConnect = sConnect & "DSN=" & cboDSNList.Text & ";"
    sConnect = sConnect & "UID=" & txtUID.Text & ";"
    sConnect = sConnect & "PWD=" & txtPWD.Text & ";"

    pobjConnection.Provider = "MSDASQL"
    pobjConnection.ConnectionString = sConnect
    pobjConnection.Open

    Connect = True
    Exit Function

Exception:
    Call mclsMsg.ShowError(Err.Description)

    Err.Clear

End Function


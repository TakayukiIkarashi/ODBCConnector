VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmSQL 
   ClientHeight    =   4260
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5610
   BeginProperty Font 
      Name            =   "Meiryo UI"
      Size            =   9.75
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSQL.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4260
   ScaleWidth      =   5610
   Begin RichTextLib.RichTextBox txtQry 
      Height          =   2055
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   3625
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      RightMargin     =   2.00000e5
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmSQL.frx":058A
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'なし
      FillColor       =   &H00808080&
      Height          =   105
      Left            =   750
      ScaleHeight     =   45.722
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   47580
      TabIndex        =   1
      Top             =   3630
      Visible         =   0   'False
      Width           =   4575
   End
   Begin MSDataGridLib.DataGrid grdQry 
      Height          =   1305
      Left            =   0
      TabIndex        =   0
      Top             =   2100
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2302
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Meiryo UI"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Meiryo UI"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Image imgSplitter 
      Height          =   135
      Left            =   -30
      MousePointer    =   7  'ｻｲｽﾞ(上下)
      Top             =   1980
      Width           =   4530
   End
End
Attribute VB_Name = "frmSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------
'   定数定義
'------------------------------
'スプリッタ 操作インデックス
Private Const SPLIT_IMG         As Integer = 1
Private Const SPLIT_PIC         As Integer = 2

Private Const SPLIT_LIMIT       As Long = 500


'------------------------------
'   変数定義
'------------------------------
Private mobjRecordSet           As Object

Private mbMoving                As Boolean

'------------------------------------------------------------
' Property Get RecordSet
'------------------------------------------------------------
Public Property Get RecordSet()

    Set RecordSet = mobjRecordSet

End Property

'------------------------------------------------------------
' Event Form_Load()
'------------------------------------------------------------
Private Sub Form_Load()

    Set mobjRecordSet = CreateObject("ADODB.Recordset")

End Sub

'------------------------------------------------------------
' Event Form_QueryUnload()
'------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call CloseRecordSet(mobjRecordSet)

    pintFrmSQLCnt = pintFrmSQLCnt - 1
    If (pintFrmSQLCnt < 0) Then
        pintFrmSQLCnt = 0
    End If

    frmMain.EnabledControls

End Sub

'------------------------------------------------------------
' Event Form_Resize()
'------------------------------------------------------------
Private Sub Form_Resize()

    txtQry.Width = Me.ScaleWidth
    grdQry.Width = txtQry.Width

    Call SizeControls(imgSplitter.Top, SPLIT_IMG)

End Sub

'------------------------------------------------------------
' Event txtQry_Change()
'------------------------------------------------------------
Private Sub txtQry_Change()
    Dim clsColorEdit As New clsEditColor

    LockWindowUpdate txtQry.hwnd

    Dim lngPos As Long
    lngPos = txtQry.SelStart

    clsColorEdit.TabSpaceNum = 4
    txtQry.TextRTF = clsColorEdit.CColor(txtQry.TextRTF)

    txtQry.SelStart = lngPos

    LockWindowUpdate 0

    Set clsColorEdit = Nothing
End Sub

'------------------------------------------------------------
' Event txtQry_OLEDragDrop()
'------------------------------------------------------------
Private Sub txtQry_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strDir                          As String
    Dim strFile                         As String
    Dim intFileNo                       As Integer
    Dim strBuf                          As String
    Dim strText                         As String

    On Error Resume Next

    strDir = Data.Files(1)

    strFile = Dir(strDir)
    If strFile = vbNullString Then
        Exit Sub
    End If

    intFileNo = FreeFile
    Open strDir For Input Access Read As intFileNo

    Do
        If EOF(intFileNo) Then
            Exit Do
        End If

        Line Input #intFileNo, strBuf

        strText = strText & strBuf & vbCrLf
    Loop

    Close intFileNo

    txtQry.Text = strText

    Err.Clear

End Sub

'------------------------------------------------------------
' Event imgSplitter_MouseDown()
'------------------------------------------------------------
Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With imgSplitter
        picSplitter.Move .Left, .Top, .Width - 20, .Height \ 2
    End With

    picSplitter.Visible = True

    mbMoving = True

End Sub

'------------------------------------------------------------
' Event imgSplitter_MouseMove()
'------------------------------------------------------------
Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single

    If mbMoving Then
        sglPos = Y + imgSplitter.Top

        If sglPos < SPLIT_LIMIT Then
            picSplitter.Top = SPLIT_LIMIT
        ElseIf sglPos > Me.Height - SPLIT_LIMIT Then
            picSplitter.Top = Me.Height - SPLIT_LIMIT
        Else
            picSplitter.Top = sglPos
        End If
    End If

End Sub

'------------------------------------------------------------
' Event imgSplitter_MouseUp()
'------------------------------------------------------------
Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call SizeControls(picSplitter.Top, SPLIT_PIC)

    picSplitter.Visible = False

    mbMoving = False

End Sub

'------------------------------------------------------------
' Sub SizeControls()
'
' スプリッタ位置変更
'
' In     : [Y]         - スプリッタ Y座標
'        : [vlngSplit] - スプリッタ 操作インデックス
' Out    : なし
' Return : なし
'------------------------------------------------------------
Private Sub SizeControls(ByVal Y As Single _
                       , ByVal vlngSplit As Long)

    On Error Resume Next

    If Y < 1500 Then Y = 1500
    If Y > (Me.Height - 3000) Then Y = Me.Height - 3000

    Select Case vlngSplit
        Case SPLIT_IMG

        Case SPLIT_PIC
            txtQry.Height = Y
            imgSplitter.Top = Y
    End Select

    txtQry.Top = 0
    grdQry.Top = txtQry.Top + txtQry.Height + 50
    grdQry.Height = Me.ScaleHeight - (txtQry.Height + 80)

    imgSplitter.Left = txtQry.Left
    imgSplitter.Width = txtQry.Width

    Err.Clear

End Sub

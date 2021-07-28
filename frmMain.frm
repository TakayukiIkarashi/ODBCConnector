VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   ClientHeight    =   4860
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7275
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows の既定値
   Begin MSComctlLib.ImageList imgGrid 
      Left            =   5160
      Top             =   1140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":115C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C90
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":222A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27C4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmnDlg 
      Left            =   5160
      Top             =   1740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgTlb 
      Left            =   5160
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2EB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3012
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":35AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B46
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":40E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":423A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":47D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4D6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5308
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":58A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5E3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":63D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6970
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6F0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":74A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  '下揃え
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   4545
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5292
            MinWidth        =   5292
            Key             =   "MESSAGE"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
            Picture         =   "frmMain.frx":7A3E
            Key             =   "DSN"
         EndProperty
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
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  '上揃え
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgTlb"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CONNECT"
            Object.ToolTipText     =   "データ ソース接続"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NEW"
            Object.ToolTipText     =   "新規作成"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OPEN"
            Object.ToolTipText     =   "SQL ファイルを開く"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SAVE"
            Object.ToolTipText     =   "SQL ファイルを保存"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EXEC"
            Object.ToolTipText     =   "実行"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ARRANGE"
            Object.ToolTipText     =   "整形"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TIMEOUT"
            Object.ToolTipText     =   "クエリ タイムアウト"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SEARCH"
            Object.ToolTipText     =   "検索"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EXCEL"
            Object.ToolTipText     =   "エクセルに貼り付け"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CASCADE"
            Object.ToolTipText     =   "重ねて表示"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "HORIZONTAL"
            Object.ToolTipText     =   "上下に並べて表示"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "VERTICAL"
            Object.ToolTipText     =   "左右に並べて表示"
            ImageIndex      =   16
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbGrid 
      Align           =   2  '下揃え
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   4125
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "imgGrid"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "START"
            Object.ToolTipText     =   "最初のレコード"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DBL_BACK"
            Object.ToolTipText     =   "10件戻る"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SGL_BACK"
            Object.ToolTipText     =   "1件戻る"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SGL_NEXT"
            Object.ToolTipText     =   "1件進む"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DBL_NEXT"
            Object.ToolTipText     =   "10件進む"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "END"
            Object.ToolTipText     =   "最後のレコード"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.ComboBox cboTable 
         BeginProperty Font 
            Name            =   "Meiryo UI"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   3
         ToolTipText     =   "テーブルの指定"
         Top             =   30
         Width           =   3495
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "ファイル(&F)"
      Begin VB.Menu mnuFileConnect 
         Caption         =   "データ ソース接続(&C)..."
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "新規作成(N)..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "SQL ファイルを開く(&O)..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "SQL ファイルを保存(&S)..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileEnd 
         Caption         =   "終了(&X)"
      End
   End
   Begin VB.Menu mnuQry 
      Caption         =   "クエリ(&Q)"
      Begin VB.Menu mnuQryExec 
         Caption         =   "実行(&R)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuQryBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQryArrange 
         Caption         =   "整形(&A)"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuQryBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQryTimeout 
         Caption         =   "クエリ タイムアウト(&T)..."
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "編集(&E)"
      Begin VB.Menu mnuEditSearch 
         Caption         =   "検索(&F)..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditResearch 
         Caption         =   "再検索(&R)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAll 
         Caption         =   "すべて選択(&A)"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuFont 
      Caption         =   "書式(&O)"
      Begin VB.Menu mnuFontFont 
         Caption         =   "フォント(&F)..."
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "ツール(&T)"
      Begin VB.Menu mnuToolExcel 
         Caption         =   "エクセルに貼り付け(&E)..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "ウィンドウ(&W)"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "新しいウィンドウを開く(&N)"
      End
      Begin VB.Menu mnuWindowBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "重ねて表示(&C)"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "上下に並べて表示(&H)"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "左右に並べて表示(&V)"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "アイコンの整列(&A)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "ヘルプ(&H)"
      Begin VB.Menu mnuHelpVersion 
         Caption         =   "バージョン(&A)..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------
'   変数定義
'------------------------------
Private mbMoving                As Boolean
Private mclsMsg                 As New clsMsg                   'メッセージ出力クラス

'------------------------------------------------------------
' Event MDIForm_Activate()
'------------------------------------------------------------
Private Sub MDIForm_Activate()

    Call EnabledControls

End Sub

'------------------------------------------------------------
' Event MDIForm_Load()
'------------------------------------------------------------
Private Sub MDIForm_Load()
    Dim strRet                  As String

    On Error Resume Next

    Me.Caption = App.Title

    With ptypWindowMinPoint
        .lngX = Me.Width \ Screen.TwipsPerPixelX
        .lngY = Me.Height \ Screen.TwipsPerPixelX
    End With
    plngOriginalhWnd = Me.hwnd
    plnglpOriginalWndProc = SetWindowLong(plngOriginalhWnd, GWL_WNDPROC, AddressOf LimitWindowSizeSubProc)

    strRet = GetSetting(App.Title, SEC_FORM_MAIN, KEY_TOP, "0")
    If IsNumeric(strRet) Then
        Top = CLng(strRet)
    Else
        Top = 0
    End If
    strRet = GetSetting(App.Title, SEC_FORM_MAIN, KEY_LEFT, "0")
    If IsNumeric(strRet) Then
        Left = CLng(strRet)
    Else
        Left = 0
    End If
    strRet = GetSetting(App.Title, SEC_FORM_MAIN, KEY_HEIGHT, "6000")
    If IsNumeric(strRet) Then
        Height = CLng(strRet)
    Else
        Height = 0
    End If
    strRet = GetSetting(App.Title, SEC_FORM_MAIN, KEY_WIDTH, "8000")
    If IsNumeric(strRet) Then
        Width = CLng(strRet)
    Else
        Width = 0
    End If
    strRet = GetSetting(App.Title, SEC_FORM_MAIN, KEY_STATE, "0")
    If IsNumeric(strRet) Then
        WindowState = CLng(strRet)
    Else
        WindowState = 0
    End If

    strRet = GetSetting(App.Title, SEC_QUERY, KEY_TIMEOUT, "30")
    If IsNumeric(strRet) Then
        pobjCommand.CommandTimeout = CLng(strRet)
    Else
        pobjCommand.CommandTimeout = 30
    End If

    Err.Clear

End Sub

'------------------------------------------------------------
' Event MDIForm_QueryUnload()
'------------------------------------------------------------
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    On Error Resume Next

    Call SetWindowLong(plngOriginalhWnd, GWL_WNDPROC, plnglpOriginalWndProc)

    If (WindowState = 0) Then
        Call SaveSetting(App.Title, SEC_FORM_MAIN, KEY_TOP, CStr(Top))
        Call SaveSetting(App.Title, SEC_FORM_MAIN, KEY_LEFT, CStr(Left))
        Call SaveSetting(App.Title, SEC_FORM_MAIN, KEY_HEIGHT, CStr(Height))
        Call SaveSetting(App.Title, SEC_FORM_MAIN, KEY_WIDTH, CStr(Width))
    End If

    Call SaveSetting(App.Title, SEC_FORM_MAIN, KEY_STATE, CStr(WindowState))
    Call SaveSetting(App.Title, SEC_QUERY, KEY_TIMEOUT, CStr(pobjCommand.CommandTimeout))

    Err.Clear

End Sub

'------------------------------------------------------------
' Event mnuFileConnect_Click()
'------------------------------------------------------------
Private Sub mnuFileConnect_Click()
    Dim strCon                  As String

    On Error Resume Next

    stbMain.Panels(STB_MESSAGE).Text = vbNullString

    Call CloseRecordSet(ActiveForm.RecordSet)

    strCon = pobjConnection.ConnectionString

    frmConnect.Show vbModal

    stbMain.Panels(STB_DSN).Text = pstrDSN

    If (strCon <> pobjConnection.ConnectionString) Then
        Call SetCboTable
    End If

    Err.Clear

End Sub

'------------------------------------------------------------
' Event mnuFileNew_Click()
'------------------------------------------------------------
Private Sub mnuFileNew_Click()

    Call LoadNewDoc

End Sub

'------------------------------------------------------------
' Event mnuFileOpen_Click()
'------------------------------------------------------------
Private Sub mnuFileOpen_Click()
    Dim strFileNm               As String
    Dim strDefFileNm            As String
    Dim blnRet                  As Boolean

    Const SEC_OPEN_SQL          As String = "OPEN_SQL"
    Const KEY_OPEN_SQL          As String = "OPEN_SQL"

    strDefFileNm = GetSetting(App.Title, SEC_OPEN_SQL, KEY_OPEN_SQL, vbNullString)
    blnRet = OpenSqlOpenDialog(cmnDlg, strFileNm, strDefFileNm)
    If (blnRet = False) Then
        Exit Sub
    End If

    If (pintFrmSQLCnt < 1) Then
        Call LoadNewDoc
    End If

    With ActiveForm
        .Caption = strFileNm
        .txtQry.Text = LoadQryFile(strFileNm)
    End With

    Call SaveSetting(App.Title, SEC_OPEN_SQL, KEY_OPEN_SQL, strFileNm)

End Sub

'------------------------------------------------------------
' Event mnuFileSave_Click()
'------------------------------------------------------------
Private Sub mnuFileSave_Click()
    Dim strRet                  As String
    Dim strFileNm               As String
    Dim strDefFileNm            As String
    Dim blnRet                  As Boolean

    Const SEC_SAVE_SQL          As String = "SAVE_SQL"
    Const KEY_SAVE_SQL          As String = "SAVE_SQL"

    strRet = LeftB(ActiveForm.Caption, 4)
    If (strRet = DEF_CAPTION) Then
        strDefFileNm = GetSetting(App.Title, SEC_SAVE_SQL, KEY_SAVE_SQL, vbNullString)
    Else
        strDefFileNm = ActiveForm.Caption
    End If
    blnRet = OpenSqlSaveDialog(cmnDlg, strFileNm, strDefFileNm)
    If (blnRet = False) Then
        Exit Sub
    End If

    Call SaveQryFile(strFileNm, ActiveForm.txtQry.Text)

    ActiveForm.Caption = strFileNm

    Call SaveSetting(App.Title, SEC_SAVE_SQL, KEY_SAVE_SQL, strFileNm)

End Sub

'------------------------------------------------------------
' Event mnuFileEnd_Click()
'------------------------------------------------------------
Private Sub mnuFileEnd_Click()

    Unload Me

End Sub

'------------------------------------------------------------
' Event mnuEditSearch_Click()
'------------------------------------------------------------
Private Sub mnuEditSearch_Click()
    Dim lngStart                    As Long

    frmSearch.Show vbModal

    lngStart = ActiveForm.txtQry.SelStart
    If (lngStart < 1) Then
        lngStart = 1
    End If

    Call SearchWord(ActiveForm, ActiveForm.txtQry, lngStart, pstrSearchWord, pblnUpperConv)

End Sub

'------------------------------------------------------------
' Event mnuEditResearch_Click()
'------------------------------------------------------------
Private Sub mnuEditResearch_Click()
    Dim lngStart                    As Long

    If (pstrSearchWord = vbNullString) Then
        frmSearch.Show vbModal
    End If

    lngStart = ActiveForm.txtQry.SelStart + ActiveForm.txtQry.SelLength
    If (lngStart < 1) Then
        lngStart = 1
    End If

    Call SearchWord(ActiveForm, ActiveForm.txtQry, lngStart, pstrSearchWord, pblnUpperConv)

End Sub

'------------------------------------------------------------
' Event mnuEditAll_Click()
'------------------------------------------------------------
Private Sub mnuEditAll_Click()

    With ActiveForm.txtQry
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

'------------------------------------------------------------
' Event mnuFontFont_Click()
'------------------------------------------------------------
Private Sub mnuFontFont_Click()

    Call ShowFontDialog(cmnDlg, Me, ActiveForm.txtQry, ActiveForm.grdQry)

    With ActiveForm.txtQry
        Call SaveSetting(App.Title, SEC_FONT, KEY_SIZE, CStr(.Font.Size))
        Call SaveSetting(App.Title, SEC_FONT, KEY_NAME, .Font.Name)
        Call SaveSetting(App.Title, SEC_FONT, KEY_BOLD, CStr(.Font.Bold))
        Call SaveSetting(App.Title, SEC_FONT, KEY_ITALIC, CStr(.Font.Italic))
    End With

End Sub

'------------------------------------------------------------
' Event mnuQryExec_Click()
'------------------------------------------------------------
Private Sub mnuQryExec_Click()
    Dim blnRet                  As Boolean

    Const STB_MSG_RUN           As String = "クエリ バッチ実行中･･･"
    Const STB_MSG_END           As String = "クエリ バッチが完了しました。"
    Const STB_MSG_ERR           As String = "クエリ バッチはエラーで完了しました。"

    On Error GoTo Exception

    stbMain.Panels(STB_MESSAGE).Text = STB_MSG_RUN

    blnRet = ExecQueryText()
    If (blnRet = False) Then
        stbMain.Panels(STB_MESSAGE).Text = STB_MSG_ERR
        Exit Sub
    End If

    Set ActiveForm.grdQry.DataSource = ActiveForm.RecordSet

    stbMain.Panels(STB_MESSAGE).Text = STB_MSG_END

    Exit Sub

Exception:
    Call mclsMsg.ShowError(Err.Description)
    stbMain.Panels(STB_MESSAGE).Text = STB_MSG_ERR

    Err.Clear

End Sub

'------------------------------------------------------------
' Event mnuQryArrange_Click()
'------------------------------------------------------------
Private Sub mnuQryArrange_Click()
    Dim clsSqlArrange   As New clsArrangeQuery
    Dim strSQL          As String

    strSQL = clsSqlArrange.CArrange(ActiveForm.txtQry.Text)

    ActiveForm.txtQry.Text = strSQL

End Sub

'------------------------------------------------------------
' Event mnuQryTimeout_Click()
'------------------------------------------------------------
Private Sub mnuQryTimeout_Click()

    frmTimeout.Show vbModal

End Sub

'------------------------------------------------------------
' Event mnuToolExcel_Click()
'------------------------------------------------------------
Private Sub mnuToolExcel_Click()
    Dim blnRet                  As Boolean

    blnRet = ExecQueryText()
    If (blnRet = False) Then Exit Sub

    Call ExcelOut(ActiveForm.RecordSet)

End Sub

'------------------------------------------------------------
' Event mnuWindowNewWindow_Click()
'------------------------------------------------------------
Private Sub mnuWindowNewWindow_Click()

    Call LoadNewDoc

End Sub

'------------------------------------------------------------
' Event mnuWindowCascade_Click()
'------------------------------------------------------------
Private Sub mnuWindowCascade_Click()

    Me.Arrange vbCascade

End Sub

'------------------------------------------------------------
' Event mnuWindowTileHorizontal_Click()
'------------------------------------------------------------
Private Sub mnuWindowTileHorizontal_Click()

    Me.Arrange vbTileHorizontal

End Sub

'------------------------------------------------------------
' Event mnuWindowTileVertical_Click()
'------------------------------------------------------------
Private Sub mnuWindowTileVertical_Click()

    Me.Arrange vbTileVertical

End Sub

'------------------------------------------------------------
' Event mnuWindowArrangeIcons_Click()
'------------------------------------------------------------
Private Sub mnuWindowArrangeIcons_Click()

    Me.Arrange vbArrangeIcons

End Sub

'------------------------------------------------------------
' Event mnuHelpVersion_Click()
'------------------------------------------------------------
Private Sub mnuHelpVersion_Click()

    frmAbout.Show vbModal

End Sub

'------------------------------------------------------------
' Event tlbMain_ButtonClick()
'------------------------------------------------------------
Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case TLB_CONNECT
            Call mnuFileConnect_Click
        Case TLB_NEW
            Call mnuFileNew_Click
        Case TLB_OPEN
            Call mnuFileOpen_Click
        Case TLB_SAVE
            Call mnuFileSave_Click

        Case TLB_EXEC
            Call mnuQryExec_Click
        Case TLB_ARRANGE
            Call mnuQryArrange_Click
        Case TLB_TIMEOUT
            Call mnuQryTimeout_Click

        Case TLB_SEARCH
            Call mnuEditSearch_Click

        Case TLB_EXCEL
            Call mnuToolExcel_Click

        Case TLB_CASCADE
            Call mnuWindowCascade_Click
        Case TLB_HORIZONTAL
            Call mnuWindowTileHorizontal_Click
        Case TLB_VERTICAL
            Call mnuWindowTileVertical_Click
    End Select

End Sub

'------------------------------------------------------------
' Event tlbGrid_ButtonClick()
'------------------------------------------------------------
Private Sub tlbGrid_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim intCnt                  As Integer

    On Error Resume Next

    Select Case Button.Key
        Case TLB_START
            ActiveForm.RecordSet.MoveFirst

        Case TLB_DBL_BACK
            For intCnt = 0 To 9
                ActiveForm.RecordSet.MovePrevious
            Next

        Case TLB_SGL_BACK
            ActiveForm.RecordSet.MovePrevious

        Case TLB_SGL_NEXT
            ActiveForm.RecordSet.MoveNext

        Case TLB_DBL_NEXT
            For intCnt = 0 To 9
                ActiveForm.RecordSet.MoveNext
            Next

        Case TLB_END
            ActiveForm.RecordSet.MoveLast

    End Select

    Err.Clear

End Sub

'------------------------------------------------------------
' Event cboTable_Click()
'------------------------------------------------------------
Private Sub cboTable_Click()
    Dim lngPos                  As Long

    If (pintFrmSQLCnt < 1) Then
        Call LoadNewDoc
    End If

    With ActiveForm.txtQry
        .SetFocus
        .Text = "SELECT * FROM " & cboTable.Text
        .SelStart = 0
    End With

End Sub

'------------------------------------------------------------
' Function ExecQueryText()
'
' SQL を実行
'
' In     : なし
' Out    : なし
' Return : 正常終了ならTrue、そうでなければFalse
'------------------------------------------------------------
Private Function ExecQueryText() As Boolean
    Dim strSQL                  As String
    Dim strQry()                As String
    Dim i                       As Integer

    On Error GoTo Exception

    ExecQueryText = False

    If (pobjConnection.State = 0) Then
        Call mclsMsg.ShowMessage("データベースに接続されていません。")
        Exit Function
    End If

    If (ActiveForm.txtQry.SelLength > 0) Then
        strSQL = ActiveForm.txtQry.SelText
    Else
        strSQL = ActiveForm.txtQry.Text
    End If

    Call SeparateSQL(strSQL, strQry())

    Screen.MousePointer = vbHourglass
    For i = 0 To UBound(strQry)
        If Not (ExecQuery(strQry(i), ActiveForm.RecordSet)) Then
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    Next i
    Screen.MousePointer = vbDefault

    ExecQueryText = True
    Exit Function

Exception:
    Call mclsMsg.ShowError(Err.Description)
    Screen.MousePointer = vbDefault

    Err.Clear

End Function

'------------------------------------------------------------
' Sub SetCboTable()
'
' テーブルコンボボックス アイテム追加
'
' In     : なし
' Out    : なし
' Return : なし
'------------------------------------------------------------
Public Sub SetCboTable()
    Dim strQry                  As String
    Dim objRs                   As Object

    On Error GoTo Finally

    If (pobjConnection.Properties(11) <> "Microsoft SQL Server") Then
        cboTable.Visible = False
        Exit Sub
    Else
        cboTable.Visible = True
    End If

    Set objRs = CreateObject("ADODB.Recordset")

    With cboTable
        .Clear

        If (pobjConnection.State <> 0) Then
            strQry = "SELECT name FROM sysobjects WHERE type = 'U' ORDER BY name"
            'ADODB.RecordSet.Open strSQL, cn, adOpenDynamic, adLockOptimistic, adCmdText
            objRs.Open strQry _
                     , pobjConnection _
                     , 2 _
                     , 3 _
                     , 1
        End If
        Do Until objRs.EOF
            .AddItem objRs.Fields("name")
            objRs.MoveNext
        Loop
    End With

Finally:
    Set objRs = Nothing

    Err.Clear

End Sub

'------------------------------------------------------------
' Sub LoadNewDoc()
'
' SQL ウィンドウ新規作成
'
' In     : なし
' Out    : なし
' Return : なし
'------------------------------------------------------------
Public Sub LoadNewDoc()
    Static lDocumentCount       As Long
    Dim frmD                    As frmSQL
    Dim strRet                  As String

    lDocumentCount = lDocumentCount + 1

    Set frmD = New frmSQL

    frmD.Caption = DEF_CAPTION & lDocumentCount

    frmD.Show

    With frmD
        strRet = GetSetting(App.Title, SEC_FONT, KEY_SIZE, CStr(DEF_FONT_SIZE))
        If IsNumeric(strRet) Then
            .txtQry.Font.Size = CLng(strRet)
        Else
            .txtQry.Font.Size = DEF_FONT_SIZE
        End If
        strRet = GetSetting(App.Title, SEC_FONT, KEY_NAME, DEF_FONT_NAME)
        .txtQry.Font.Name = strRet
        strRet = GetSetting(App.Title, SEC_FONT, KEY_BOLD, CStr(False))
        .txtQry.Font.Bold = CBool(strRet)
        strRet = GetSetting(App.Title, SEC_FONT, KEY_ITALIC, CStr(False))
        .txtQry.Font.Italic = CBool(strRet)
        .grdQry.Font.Size = .txtQry.Font.Size
        .grdQry.Font.Name = .txtQry.Font.Name
        .grdQry.Font.Bold = .txtQry.Font.Bold
        .grdQry.Font.Italic = .txtQry.Font.Italic
    End With

    pintFrmSQLCnt = pintFrmSQLCnt + 1

    Call EnabledControls

End Sub

'------------------------------------------------------------
' Sub EnabledControls()
'
' コントロールの使用可否設定
'
' In     : なし
' Out    : なし
' Return : なし
'------------------------------------------------------------
Public Sub EnabledControls()
    Dim blnEnabled              As Boolean

    blnEnabled = (pintFrmSQLCnt > 0)

    mnuFileSave.Enabled = blnEnabled
    mnuQry.Enabled = blnEnabled
    mnuEdit.Enabled = blnEnabled
    mnuFont.Enabled = blnEnabled
    mnuTool.Enabled = blnEnabled
    mnuWindowCascade.Enabled = blnEnabled
    mnuWindowTileHorizontal.Enabled = blnEnabled
    mnuWindowTileVertical.Enabled = blnEnabled
    mnuWindowArrangeIcons.Enabled = blnEnabled

    tlbMain.Buttons(TLB_SAVE).Enabled = blnEnabled
    tlbMain.Buttons(TLB_EXEC).Enabled = blnEnabled
    tlbMain.Buttons(TLB_ARRANGE).Enabled = blnEnabled
    tlbMain.Buttons(TLB_TIMEOUT).Enabled = blnEnabled
    tlbMain.Buttons(TLB_SEARCH).Enabled = blnEnabled
    tlbMain.Buttons(TLB_EXCEL).Enabled = blnEnabled
    tlbMain.Buttons(TLB_CASCADE).Enabled = blnEnabled
    tlbMain.Buttons(TLB_HORIZONTAL).Enabled = blnEnabled
    tlbMain.Buttons(TLB_VERTICAL).Enabled = blnEnabled

End Sub

'------------------------------------------------------------
' Sub SeparateSQL()
'
' 引数に指定されたSQLを"GO"文で分割
'
' In     : [vstrSQL]   - 分割するSQL
' Out    : [rstrQry()] - 構造体に格納されたクエリ
' Return : なし
'------------------------------------------------------------
Private Sub SeparateSQL(ByVal vstrSQL As String _
                      , ByRef rstrQry() As String)
    Dim strSQLTmp   As String
    Dim strSepRet() As String
    Dim strQry      As String
    Dim strUnit     As String
    Dim strTest     As String
    Dim i           As Integer
    Dim j           As Integer

    strSQLTmp = vstrSQL

    'vbLfは全てvbCrする
    strSQLTmp = Replace(strSQLTmp, vbLf, vbCr)

    '改行ｺｰﾄﾞ(vbCr)にてSQL文字列を分割する
    strSepRet = Split(strSQLTmp, vbCr)

    '変数を初期化する
    strQry = vbNullString
    j = 0

    '改行ｺｰﾄﾞ分だけﾙｰﾌﾟを繰り返し、"GO"文を検索する
    For i = 0 To UBound(strSepRet)
        '分割文字列を取得
        strUnit = strSepRet(i)

        '改行ｺｰﾄﾞを空白に変換する
        strTest = Replace(strUnit, vbCr, vbNullString)

        '分割文字列が"GO"であれば、その時点でSQLを分割する
        If (Trim(UCase(strTest)) = "GO") Then
            ReDim Preserve rstrQry(j)
            rstrQry(j) = strQry
            j = j + 1
            strQry = vbNullString
        Else
            strQry = strQry & strUnit & vbCr
        End If
    Next i

    '最後のｸｴﾘを戻り値の配列にｾｯﾄ
    ReDim Preserve rstrQry(j)
    rstrQry(j) = strQry

End Sub



Attribute VB_Name = "basMain"
Option Explicit
Option Base 0

'------------------------------
'   WinAPI関数定義
'------------------------------
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = (-4)

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const WM_GETMINMAXINFO = &H24


'------------------------------
'   構造体定義
'------------------------------
Private Type tagPOINT
    lngX                        As Long
    lngY                        As Long
End Type
Private Type MINMAXINFO
    ptReserved                  As tagPOINT
    ptMaxSize                   As tagPOINT
    ptMaxPosition               As tagPOINT
    ptMinTrackSize              As tagPOINT
    ptMaxTrackSize              As tagPOINT
End Type
Public ptypWindowMinPoint       As tagPOINT
Public plnglpOriginalWndProc    As Long
Public plngOriginalhWnd         As Long

Private Type PTEXLCELL
    lngRow                      As Long
    lngCol                      As Long
    strRange                    As String
End Type


'------------------------------
'   定数定義
'------------------------------
'レジストリ設定 キー
Public Const SEC_FORM_MAIN      As String = "FORM_MAIN"
Public Const KEY_TOP            As String = "TOP"
Public Const KEY_LEFT           As String = "LEFT"
Public Const KEY_HEIGHT         As String = "HEIGHT"
Public Const KEY_WIDTH          As String = "WIDTH"
Public Const KEY_STATE          As String = "STATE"

Public Const SEC_FONT           As String = "FONT"
Public Const KEY_SIZE           As String = "SIZE"
Public Const KEY_NAME           As String = "NAME"
Public Const KEY_BOLD           As String = "BOLD"
Public Const KEY_ITALIC         As String = "ITALIC"

Public Const SEC_QUERY          As String = "QUERY"
Public Const KEY_TIMEOUT        As String = "TIMEOUT"

Public Const SEC_SEARCH         As String = "SEARCH"
Public Const KEY_UPPER          As String = "UPPER"


'メインツールバー キー
Public Const TLB_CONNECT        As String = "CONNECT"
Public Const TLB_NEW            As String = "NEW"
Public Const TLB_OPEN           As String = "OPEN"
Public Const TLB_SAVE           As String = "SAVE"

Public Const TLB_EXEC           As String = "EXEC"
Public Const TLB_ARRANGE        As String = "ARRANGE"
Public Const TLB_TIMEOUT        As String = "TIMEOUT"

Public Const TLB_SEARCH         As String = "SEARCH"

Public Const TLB_EXCEL          As String = "EXCEL"

Public Const TLB_CASCADE        As String = "CASCADE"
Public Const TLB_HORIZONTAL     As String = "HORIZONTAL"
Public Const TLB_VERTICAL       As String = "VERTICAL"


'DataGrid ツールバー キー
Public Const TLB_START          As String = "START"
Public Const TLB_DBL_BACK       As String = "DBL_BACK"
Public Const TLB_SGL_BACK       As String = "SGL_BACK"
Public Const TLB_SGL_NEXT       As String = "SGL_NEXT"
Public Const TLB_DBL_NEXT       As String = "DBL_NEXT"
Public Const TLB_END            As String = "END"


'ステータスバー キー
Public Const STB_MESSAGE        As Integer = 1
Public Const STB_DSN            As Integer = 2


'SQLフォーム
Public Const DEF_CAPTION        As String = "無題"


'初期フォント
Public Const DEF_FONT_SIZE      As Integer = 10
Public Const DEF_FONT_NAME      As String = "Meiryo UI"


'ディレクトリ操作
Private Const INIT_DIR          As String = "\"


'------------------------------
'   変数定義
'------------------------------
Public pobjConnection           As Object
Public pobjCommand              As Object
Public pstrDSN                  As String

Public pintFrmSQLCnt            As Integer

Public pstrSearchWord           As String
Public pblnUpperConv            As Boolean

Private mclsMsg                 As New clsMsg                   'メッセージ出力クラス

'------------------------------------------------------------
' Sub Main()
'
' 起動時処理
'
' In     : なし
' Out    : なし
' Return : なし
'------------------------------------------------------------
Sub Main()

    On Error GoTo Exception

    Set pobjConnection = CreateObject("ADODB.Connection")
    Set pobjCommand = CreateObject("ADODB.Command")

    pstrDSN = vbNullString
    pintFrmSQLCnt = 0

    pstrSearchWord = vbNullString
    pblnUpperConv = False

    frmMain.Show

    frmConnect.Show vbModal

    With frmMain
        .LoadNewDoc
        .SetCboTable
        .stbMain.Panels(STB_DSN).Text = pstrDSN
    End With

    Exit Sub

Exception:
    Call mclsMsg.ShowError(Err.Description)

    Err.Clear

End Sub

'------------------------------------------------------------
' Function ExecQuery()
'
' クエリの実行
'
' In     : [vstrQry] - SQL
'        : [vblnErr] - 再接続フラグ(省略可)
' Out    : [robjRS]  - ADODB.RecordSet
' Return : 正常終了ならTrue、そうでなければFalse
'------------------------------------------------------------
Public Function ExecQuery(ByVal vstrQry As String _
                        , ByRef robjRS As Object _
               , Optional ByVal vblnErr As Boolean = False) As Boolean

    ExecQuery = False

    On Error GoTo Exception

    Call CloseRecordSet(robjRS)

    If (vstrQry = vbNullString) Then
        ExecQuery = True
        Exit Function
    End If

    With pobjCommand
        .ActiveConnection = pobjConnection
        .CommandText = vstrQry
    End With

    With robjRS
        .CursorLocation = 3         'adUseClient    '2008.05.18 -s
        .CursorType = 1             'adOpenKeyset
        .LockType = 3               'adLockOptimistic

        .Open pobjCommand
    End With

    ExecQuery = True
    Exit Function

Exception:
    If (Err.Number = 3705) And (vblnErr = False) Then
        Err.Clear

        robjRS.CancelUpdate
        Call ExecQuery(vstrQry, robjRS, True)

        ExecQuery = True
        Exit Function
    End If

    Call mclsMsg.ShowError(Err.Description)

    Err.Clear

End Function

'------------------------------------------------------------
' Sub CloseRecordSet()
'
' ADODB.RecordSet を閉じる
'
' In     : [robjRS] - ADODB.RecordSet
' Out    : なし
' Return : なし
'------------------------------------------------------------
Public Sub CloseRecordSet(ByRef robjRS As Object)

    On Error Resume Next

    pobjConnection.CommitTrans

    robjRS.Close

    Err.Clear

End Sub

'------------------------------------------------------------
' Function OpenSqlSaveDialog()
'
' 「SQL ファイルを保存」ダイアログ ボックスを表示
'
' In     : [vobjCmnDlg]    - 共通ダイアログ コントロール
'        : [vstrDefFileNm] - 初期表示ファイル名
' Out    : [rstrFileNm]    - ダイアログで選択したファイル名
' Return : 正常終了ならTrue、そうでなければFalse
'------------------------------------------------------------
Public Function OpenSqlSaveDialog(ByVal vobjCmnDlg As CommonDialog _
                                , ByRef rstrFileNm As String _
                                , ByVal vstrDefFileNm As String) As Boolean
    Dim intRet                  As Integer

    On Error GoTo Exception

    OpenSqlSaveDialog = False

    rstrFileNm = vbNullString

    With vobjCmnDlg
        .CancelError = True
        .Flags = cdlOFNCreatePrompt Or _
                 cdlOFNHideReadOnly Or _
                 cdlOFNNoReadOnlyReturn Or _
                 cdlOFNOverwritePrompt
        .Filter = "SQLファイル (*.sql)|*.sql|すべてのファイル (*.*)|*.*"
        .FilterIndex = 1
        .InitDir = INIT_DIR
        .FileName = vstrDefFileNm
    End With

    On Error Resume Next

    vobjCmnDlg.ShowSave

    If (Err.Number = cdlCancel) Then
        Exit Function
    Else
        If (Err.Number <> 0) Then
            Call mclsMsg.ShowError(Err.Description)
            Exit Function
        End If
    End If

    vobjCmnDlg.Parent.Refresh

    Err.Clear
    On Error GoTo Exception

    rstrFileNm = vobjCmnDlg.FileName

    OpenSqlSaveDialog = True

    Exit Function

Exception:
    Call mclsMsg.ShowError(Err.Description)

    Err.Clear

End Function

'------------------------------------------------------------
' Function OpenSqlOpenDialog()
'
' 「SQL ファイルを開く」ダイアログ ボックスを表示
'
' In     : [vobjCmnDlg]    - 共通ダイアログ コントロール
'        : [vstrDefFileNm] - 初期表示ファイル名
' Out    : [rstrFileNm]    - ダイアログで選択したファイル名
' Return : 正常終了ならTrue、そうでなければFalse
'------------------------------------------------------------
Public Function OpenSqlOpenDialog(ByVal vobjCmnDlg As CommonDialog _
                                , ByRef rstrFileNm As String _
                                , ByVal vstrDefFileNm As String) As Boolean
    Dim intRet                  As Integer

    On Error GoTo Exception

    OpenSqlOpenDialog = False

    rstrFileNm = vbNullString

    With vobjCmnDlg
        .CancelError = True
        .Flags = cdlOFNCreatePrompt Or _
                 cdlOFNHideReadOnly Or _
                 cdlOFNNoReadOnlyReturn Or _
                 cdlOFNOverwritePrompt
        .Filter = "SQLファイル (*.sql)|*.sql|すべてのファイル (*.*)|*.*"
        .FilterIndex = 1
        .InitDir = INIT_DIR
        .FileName = vstrDefFileNm
    End With

    On Error Resume Next

    vobjCmnDlg.ShowOpen

    If (Err.Number = cdlCancel) Then
        Exit Function
    Else
        If (Err.Number <> 0) Then
            Call mclsMsg.ShowError(Err.Description)
            Exit Function
        End If
    End If

    vobjCmnDlg.Parent.Refresh

    Err.Clear
    On Error GoTo Exception

    rstrFileNm = vobjCmnDlg.FileName

    OpenSqlOpenDialog = True

    Exit Function

Exception:
    Call mclsMsg.ShowError(Err.Description)

    Err.Clear

End Function

'------------------------------------------------------------
' Function ShowFontDialog()
'
' 「フォントを指定」ダイアログ ボックスを表示
'
' In     : [vobjCmnDlg]   - 共通ダイアログ コントロール
' Out    : [robjForm]     - フォーム オブジェクト
'          [robjText]     - テキストボックス オブジェクト
'          [robjDataGrid] - データグリッド オブジェクト
' Return : 正常終了ならTrue、そうでなければFalse
'------------------------------------------------------------
Public Function ShowFontDialog(ByVal vobjCmnDlg As CommonDialog _
                             , ByRef robjForm As Form _
                             , ByRef robjText As RichTextBox _
                    , Optional ByRef robjDataGrid As DataGrid = Nothing) As Boolean
    Dim intRet                  As Integer

    On Error GoTo Exception

    ShowFontDialog = False

    With vobjCmnDlg
        .Flags = cdlCFScreenFonts
        .FontSize = robjText.Font.Size
        .FontName = robjText.Font.Name
        .FontBold = robjText.Font.Bold
        .FontItalic = robjText.Font.Italic
    End With

    On Error Resume Next

    vobjCmnDlg.ShowFont

    If (Err.Number = cdlCancel) Then
        Exit Function
    Else
        If (Err.Number <> 0) Then
            Call mclsMsg.ShowError(Err.Description)
            Exit Function
        End If
    End If

    vobjCmnDlg.Parent.Refresh

    Err.Clear
    On Error GoTo Exception

    With vobjCmnDlg
        robjText.Font.Size = .FontSize
        robjText.Font.Name = .FontName
        robjText.Font.Bold = .FontBold
        robjText.Font.Italic = .FontItalic
        If Not (robjDataGrid Is Nothing) Then
            robjDataGrid.Font.Size = .FontSize
            robjDataGrid.Font.Name = .FontName
            robjDataGrid.Font.Bold = .FontBold
            robjDataGrid.Font.Italic = .FontItalic
        End If
    End With

    ShowFontDialog = True

    Exit Function

Exception:
    Call mclsMsg.ShowError(Err.Description)

    Err.Clear

End Function

'------------------------------------------------------------
' Function LoadQryFile()
'
' SQL ファイルを開く
'
' In     : [vstrFile] - テキストファイル パス
' Out    : なし
' Return : 読み込んだSQL 文字列
'------------------------------------------------------------
Public Function LoadQryFile(ByVal vstrFile As String) As String
    Dim intFile                 As Integer
    Dim strLine                 As String
    Dim strQry                  As String

    On Error GoTo Exception

    LoadQryFile = vbNullString

    If (vstrFile = vbNullString) Then Exit Function
    If Dir(vstrFile) = vbNullString Then Exit Function

    intFile = FreeFile
    strQry = vbNullString

    Open vstrFile For Input As intFile
    While Not EOF(intFile)
        Line Input #intFile, strLine
        strQry = strQry & strLine & vbCrLf
    Wend
    Close intFile

    LoadQryFile = strQry
    Exit Function

Exception:
    Call mclsMsg.ShowError(Err.Description)

    Err.Clear

End Function

'------------------------------------------------------------
' Sub SaveQryFile
'
' SQL ファイルを保存
'
' In     : [vstrFile] - テキストファイル パス
'        : [vstrQry]  - SQL 文字列
' Out    : なし
' Return : なし
'------------------------------------------------------------
Public Sub SaveQryFile(ByVal vstrFile As String _
                     , ByVal vstrQry As String)
    Dim intFile                 As Integer
    Dim strQry                  As String

    On Error GoTo Exception

    If (vstrFile = vbNullString) Then Exit Sub
    If (vstrQry = vbNullString) Then Exit Sub

    intFile = FreeFile

    Open vstrFile For Output As intFile
    Print #intFile, vstrQry
    Close intFile

    Exit Sub

Exception:
    Call mclsMsg.ShowError(Err.Description)

    Err.Clear

End Sub

'------------------------------------------------------------
' Sub ExcelOut
'
' 「エクセルに貼り付け」処理
'
' In     : なし
' Out    : [robjRS] - ADODB.RecordSet
' Return : なし
'------------------------------------------------------------
Public Sub ExcelOut(ByRef robjRS As Object)
    Dim objExcelApp             As Object
    Dim objWorkBook             As Object
    Dim objWorkSheet            As Object
    Dim objRange                As Object
    Dim typCell                 As PTEXLCELL
    Dim ilngIdx                 As Long

    On Error GoTo Exception

    Set objExcelApp = CreateObject("Excel.Application")

    objExcelApp.Interactive = False

    Set objWorkBook = objExcelApp.Workbooks.Add
    Set objWorkSheet = objWorkBook.ActiveSheet
    Set objRange = objWorkSheet.Range("A1")

    For ilngIdx = 0 To robjRS.Fields.Count - 1
        objRange.Parent.Columns(ilngIdx + 1).NumberFormat = "@"
        objRange.Value = robjRS.Fields(ilngIdx).Name
        Set objRange = objRange.Offset(0, 1)
    Next ilngIdx

    typCell.strRange = "A2"
    typCell.lngCol = 1
    typCell.lngRow = 2

    Call CopyResultSet(robjRS, objWorkSheet, typCell)

    objExcelApp.Interactive = True
    objExcelApp.Visible = True
    objWorkSheet.Activate

    Set objExcelApp = Nothing
    Exit Sub

Exception:
    Call mclsMsg.ShowError(Err.Description)

    On Error Resume Next

    objExcelApp.Interactive = True
    objExcelApp.Visible = True
    Set objExcelApp = Nothing

    Err.Clear

End Sub

'------------------------------------------------------------
' Sub CopyResultSet
'
' レコードセットをエクセルに出力
'
' In     : なし
' Out    : [robjRS]   - ADODB.RecordSet
'        : [robjWS]   - エクセル ワークシート
'        : [vtypCell] - エクセル セル
' Return : なし
'------------------------------------------------------------
Private Sub CopyResultSet(ByRef robjRS As Object _
                        , ByRef robjWS As Object _
                        , ByRef rtypCell As PTEXLCELL)
    Dim varSomeArray()          As Variant
    Dim lngArrayCols            As Long
    Dim ilngRow                 As Long
    Dim ilngCol                 As Long
    Dim ilngIdx                 As Long

    Const clngArrayBuff         As Long = 100

    On Error GoTo Exception

    lngArrayCols = robjRS.Fields.Count

    With robjRS
        Do While (.EOF = False)
            ReDim varSomeArray(clngArrayBuff - 1, lngArrayCols - 1)
            For ilngRow = 0 To clngArrayBuff - 1
                ilngCol = 0
                For ilngIdx = 0 To .Fields.Count - 1
                    varSomeArray(ilngRow, ilngCol) = .Fields(ilngIdx).Value
                    If IsNull(varSomeArray(ilngRow, ilngCol)) Then
                        varSomeArray(ilngRow, ilngCol) = vbNullString
                    End If
                    ilngCol = ilngCol + 1
                Next
                .MoveNext

                If (.EOF = True) Then
                    Exit For
                End If
            Next ilngRow

            robjWS.Range(robjWS.Cells(rtypCell.lngRow _
                                    , rtypCell.lngCol) _
                       , robjWS.Cells(rtypCell.lngRow + clngArrayBuff - 1 _
                                    , rtypCell.lngCol + lngArrayCols - 1)).Value = varSomeArray
            rtypCell.lngRow = rtypCell.lngRow + clngArrayBuff
        Loop
        .Close
    End With

    Exit Sub

Exception:
'    Call mclsMsg.ShowError(Err.Description)
    Debug.Print Err.Description

    Err.Clear

End Sub

'------------------------------------------------------------
' Function LimitWindowSizeSubProc
'
' フォームの最小幅を設定
'
' In     : [hWnd]   - ハンドル
'        : [uMsg]   - メッセージ
'        : [wParam] - パラメータ1
'        : [lParam] - パラメータ2
' Out    : なし
' Return : Win32 API戻り値
'------------------------------------------------------------
Public Function LimitWindowSizeSubProc(ByVal hwnd As Long _
                                     , ByVal uMsg As Long _
                                     , ByVal wParam As Long _
                                     , ByVal lParam As Long) As Long
    Dim typWinSize As MINMAXINFO

    If (uMsg = WM_GETMINMAXINFO) Then
        MoveMemory typWinSize, ByVal lParam, Len(typWinSize)
        LSet typWinSize.ptMinTrackSize = ptypWindowMinPoint
        MoveMemory ByVal lParam, typWinSize, Len(typWinSize)

        LimitWindowSizeSubProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
    Else
        LimitWindowSizeSubProc = CallWindowProc(plnglpOriginalWndProc, hwnd, uMsg, wParam, lParam)
    End If

End Function

'------------------------------------------------------------
' Sub SearchWord()
'
' キーワード検索処理
'
' In     : [vlngStart]     - 検索開始位置
'          [vstrWord]      - 検索文字列
'          [vblnUpperConv] - 大文字変換するしない
' Out    : [robjForm]      - Form
'          [robjTextBox]   - RichTextBox
' Return : なし
'------------------------------------------------------------
Public Sub SearchWord(ByRef robjForm As Form _
                    , ByRef robjTextBox As RichTextBox _
                    , ByVal vlngStart As Long _
                    , ByVal vstrWord As String _
                    , ByVal vblnUpperConv As Boolean)
    Dim strText                         As String
    Dim strWord                         As String
    Dim lngFind                         As Long

    On Error Resume Next

    If (vblnUpperConv = True) Then
        strWord = StrConv(vstrWord, vbUpperCase)
        strText = StrConv(robjTextBox.Text, vbUpperCase)
    Else
        strWord = vstrWord
        strText = robjTextBox.Text
    End If

    lngFind = InStr(vlngStart, strText, strWord)
    If (lngFind > 0) Then
        robjTextBox.SelStart = lngFind - 1
        robjTextBox.SelLength = Len(strWord)
    End If

    Err.Clear

End Sub


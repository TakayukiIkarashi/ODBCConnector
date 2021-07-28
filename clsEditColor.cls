VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEditColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum CommentType
    None = 0
    SingleLine = 1
    MultiLine = 2
End Enum

Private Type Keyword
    Character As String
    Color As Integer
End Type

Private Const COLOR_TABLE As String = _
    "{\colortbl ;\red0\green0\blue255;\red100\green100\blue100;\red255\green0\blue255;\red0\green150\blue0;}"

Private Const RTF_CF0 = "\cf0 "
Private Const RTF_CF1 = "\cf1 "
Private Const RTF_CF2 = "\cf2 "
Private Const RTF_CF3 = "\cf3 "
Private Const RTF_CF4 = "\cf4 "
Private Const RTF_PAR = "\par "
Private Const RTF_TAB = "\tab"

Private Const KEYWORD_FILE As String = "\KEYWORD_COLOR.CSV"

Private mintTabSpaceNum As Integer

Private Sub Class_Initialize()
    mintTabSpaceNum = 4
End Sub

Public Property Get TabSpaceNum() As Integer
    TabSpaceNum = mintTabSpaceNum
End Property

Public Property Let TabSpaceNum(ByVal vintNum As Integer)
    mintTabSpaceNum = vintNum
End Property

Public Function CColor(ByVal vstrText As String) As String
    CColor = ""

    vstrText = ReplaceBeforeChange(vstrText)

    Dim typKeyword() As Keyword
    If ReadFile(typKeyword()) = False Then
        Exit Function
    End If

    Dim strRet As String
    strRet = ""

    Call InitColor(vstrText)

    If EditColor(vstrText, typKeyword(), strRet) = False Then
        Exit Function
    End If

    CColor = strRet
End Function

Private Function ReplaceBeforeChange(ByVal vstrText As String) As String
    vstrText = Replace(vstrText, vbCrLf, "")
    vstrText = Replace(vstrText, vbCr, "")
    vstrText = Replace(vstrText, vbLf, "")
    Dim strSpace As String
    strSpace = ""
    Dim i As Integer
    For i = 1 To TabSpaceNum - 1
        strSpace = strSpace & " "
    Next i
    vstrText = Replace(vstrText, RTF_TAB, strSpace)

    ReplaceBeforeChange = vstrText
End Function

Private Function ReadFile(ByRef rtypKeyword() As Keyword) As Boolean
    ReadFile = False

    Dim strCt As String
    strCt = ""

    If GetKeyword(strCt) = False Then
        Exit Function
    End If

    If SetKeyword(strCt, rtypKeyword()) = False Then
        Exit Function
    End If

    ReadFile = True
End Function

Private Function GetKeyword(ByRef rstrCt As String) As Boolean
    GetKeyword = False

    Dim objFso As Object
    Set objFso = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next

    Dim objTs As Object
    Set objTs = objFso.OpenTextFile(App.Path & KEYWORD_FILE, 1, False)

    If Err.Number <> 0 Then
        Call ShowError(Err)
        Exit Function
    End If

    On Error GoTo 0

    rstrCt = objTs.ReadAll

    Set objTs = Nothing
    Set objFso = Nothing

    GetKeyword = True
End Function

Private Function SetKeyword(ByVal vstrCt As String, ByRef rtypKeyword() As Keyword) As Boolean
    SetKeyword = False

    On Error GoTo Exception

    Dim arrLine As Variant
    arrLine = Split(vstrCt, vbCrLf)

    Dim i As Integer
    For i = 1 To UBound(arrLine)
        Dim strLine As String
        strLine = arrLine(i)

        If strLine <> "" Then
            Dim arrData As Variant
            arrData = GetCsvValue(strLine)

            ReDim Preserve rtypKeyword(i - 1)
            rtypKeyword(i - 1).Character = arrData(0)
            rtypKeyword(i - 1).Color = CInt(arrData(1))
        End If
    Next i

    SetKeyword = True
    Exit Function

Exception:
    Call ShowError(Err)
End Function

Private Sub InitColor(ByRef rstrText As String)
    rstrText = Replace(rstrText, RTF_CF0, "")
    rstrText = Replace(rstrText, RTF_CF1, "")
    rstrText = Replace(rstrText, RTF_CF2, "")
    rstrText = Replace(rstrText, RTF_CF3, "")
    rstrText = Replace(rstrText, RTF_CF4, "")
End Sub

Private Function EditColor( _
    ByVal vstrBefore As String, _
    ByRef rtypKeyword() As Keyword, _
    ByRef rstrAfter As String _
) As Boolean
    EditColor = False

    On Error GoTo Exception

    rstrAfter = "{" & COLOR_TABLE & Right(vstrBefore, Len(vstrBefore) - 1)

    Dim i As Integer
    For i = 0 To UBound(rtypKeyword())

        Dim lngPos As Long
        lngPos = 1

        Do
            If lngPos >= Len(rstrAfter) Then
                Exit Do
            End If

            lngPos = InStr(lngPos, LCase(rstrAfter), LCase(rtypKeyword(i).Character))
            If lngPos < 1 Then
                Exit Do
            End If

            Dim lngPosNxtCh As Long
            lngPosNxtCh = lngPos + Len(rtypKeyword(i).Character)

            Dim lngPosPriCh As Long
            lngPosPriCh = lngPos - 1
            Dim strPriCh As String
            If lngPosPriCh > 0 Then
                strPriCh = Mid(rstrAfter, lngPosPriCh, 1)
            Else
                strPriCh = ""
            End If

            Dim strNxtCh As String
            strNxtCh = Mid(rstrAfter, lngPosNxtCh, 1)

            If IsSeparate(strPriCh) And IsSeparate(strNxtCh) Then
                rstrAfter = CColorKeyword(rstrAfter, rtypKeyword(i), lngPos, lngPosNxtCh)
            End If

            lngPos = lngPosNxtCh + 1
        Loop
    Next i

    '2020.05.23
    'Visual Studio Installer 1.1で作成したインストーラーだと、なぜかSpace(1)の付加が必要となる。
    'これがないと、"SELECT * "と入力した後、"FROM"の"F"を入力した時点で、"SELECT* F"となってしまう。
    rstrAfter = CColorComment(rstrAfter) & Space(1)

    EditColor = True
    Exit Function

Exception:
    Call ShowError(Err)
End Function

Private Function CColorKeyword( _
    ByVal vstrBefore As String, _
    ByRef rtypKeyword As Keyword, _
    ByVal vlngPos As Long, _
    ByRef rlngPosNxtCh As Long _
) As String

    Dim strLeft As String
    strLeft = Left(vstrBefore, vlngPos - 1)

    Dim strTarget As String
    strTarget = Mid(vstrBefore, vlngPos, Len(rtypKeyword.Character))
    strTarget = AddColor(rtypKeyword.Color, strTarget)

    Dim strRight As String
    If Len(vstrBefore) >= rlngPosNxtCh Then
        strRight = Right(vstrBefore, Len(vstrBefore) - rlngPosNxtCh + 1)
    Else
        strRight = ""
    End If

    rlngPosNxtCh = Len(strLeft & strTarget)

    CColorKeyword = strLeft & strTarget & strRight
End Function

Private Function AddColor(ByVal vintColor As Integer, ByVal vstrTarget As String) As String
    Dim strCf As String

    Select Case vintColor
        Case 1
            strCf = RTF_CF1
        Case 2
            strCf = RTF_CF2
        Case 3
            strCf = RTF_CF3
        Case Else
            strCf = RTF_CF0
    End Select

    AddColor = strCf & vstrTarget & RTF_CF0
End Function

Private Function IsSeparate(ByVal vstrChar As String) As Boolean
    If _
        (vstrChar = "") Or _
        (vstrChar = " ") Or _
        (vstrChar = "　") Or _
        (vstrChar = "\") Then
        IsSeparate = True
    Else
        IsSeparate = False
    End If
End Function

Private Function CColorComment(ByVal vstrText As String) As String
    Dim lngPos As Long
    lngPos = 1

    Do
        Dim lngCmStS, lngCmStM As Long

        lngCmStS = InStr(lngPos, vstrText, "--")
        If lngCmStS < 1 Then
            lngCmStS = Len(vstrText)
        End If

        lngCmStM = InStr(lngPos, vstrText, "/*")
        If lngCmStM < 1 Then
            lngCmStM = Len(vstrText)
        End If

        Dim intCmType As CommentType
        intCmType = None

        If lngCmStS < lngCmStM Then
            intCmType = SingleLine
        End If

        If lngCmStS > lngCmStM Then
            intCmType = MultiLine
        End If

        Dim lngCmSt, lngCmEd As Long

        Select Case intCmType
            Case SingleLine
                lngCmSt = InStr(lngPos, vstrText, "--")
                lngCmEd = InStr(lngCmSt, vstrText, RTF_PAR)
                If lngCmEd < 1 Then
                    lngCmEd = Len(vstrText)
                End If
            Case MultiLine
                lngCmSt = InStr(lngPos, vstrText, "/*")
                lngCmEd = InStr(lngCmSt, vstrText, "*/")
                If lngCmEd < 1 Then
                    lngCmEd = Len(vstrText)
                End If
            Case Else
                Exit Do
        End Select

        Dim strLeft As String
        Dim strTgt As String
        Dim strRight As String

        strLeft = Left(vstrText, lngCmSt - 1)

        Select Case intCmType
            Case SingleLine
                strTgt = Mid(vstrText, lngCmSt, lngCmEd - lngCmSt)
            Case MultiLine
                strTgt = Mid(vstrText, lngCmSt, lngCmEd - lngCmSt + 2)
        End Select
        Call InitColor(strTgt)

        If lngCmEd < Len(vstrText) Then
            strTgt = RTF_CF4 & strTgt & RTF_CF0
            Select Case intCmType
                Case SingleLine
                    strRight = Right(vstrText, Len(vstrText) - lngCmEd + 1)
                Case MultiLine
                    strRight = Right(vstrText, Len(vstrText) - lngCmEd - 1)
            End Select
        Else
            strTgt = RTF_CF4 & Left(strTgt, Len(strTgt) - 1) & RTF_CF0 & "}"
            strRight = ""
        End If

        vstrText = strLeft & strTgt & strRight

        lngPos = lngCmEd
    Loop

    CColorComment = vstrText
End Function

Private Sub ShowError(ByVal vobjErr As ErrObject)
    MsgBox CStr(vobjErr.Number) & ":" & vobjErr.Description, vbCritical + vbOKOnly, App.Title
End Sub

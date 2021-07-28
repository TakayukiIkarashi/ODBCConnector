VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArrangeQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Keyword
    Character As String
    IndentLevel As Integer
    SeparateFlag As Integer
End Type

Private Const INDENT_UNIT As Integer = 4

Private Const SEPARATEFLAG_ON As Integer = 1
Private Const SEPARATEFLAG_OFF As Integer = 0

Private Const KEYWORD_FILE As String = "\KEYWORD_ARRANGE.CSV"

Public Function CArrange(ByVal vstrQry As String) As String
    CArrange = ""

    Dim typKeyword() As Keyword
    If ReadFile(typKeyword()) = False Then
        Exit Function
    End If

    Dim strRet As String
    strRet = ""

    If EditQuery(vstrQry, typKeyword(), strRet) = False Then
        Exit Function
    End If

    CArrange = strRet
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
        strLine = Trim(arrLine(i))

        If strLine <> "" Then
            Dim arrData As Variant
            arrData = GetCsvValue(strLine)

            ReDim Preserve rtypKeyword(i - 1)
            rtypKeyword(i - 1).Character = arrData(0)
            rtypKeyword(i - 1).IndentLevel = arrData(1)
            rtypKeyword(i - 1).SeparateFlag = arrData(2)
        End If
    Next i

    SetKeyword = True
    Exit Function

Exception:
    Call ShowError(Err)
End Function

Private Function EditQuery( _
    ByVal vstrBefore As String, _
    ByRef rtypKeyword() As Keyword, _
    ByRef rstrAfter As String _
) As Boolean
    EditQuery = False

    On Error GoTo Exception

    rstrAfter = ReplaceSpecialChar(vstrBefore)

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

            If rtypKeyword(i).SeparateFlag = SEPARATEFLAG_ON Then
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
                    rstrAfter = CQueryKeyword(rstrAfter, rtypKeyword(i), lngPos, lngPosNxtCh)
                End If
            Else
                rstrAfter = CQueryKeyword(rstrAfter, rtypKeyword(i), lngPos, lngPosNxtCh)
            End If

            lngPos = lngPosNxtCh + 1
        Loop
    Next i

    rstrAfter = ReplaceSpecialCharUndo(rstrAfter)

    EditQuery = True
    Exit Function

Exception:
    Call ShowError(Err)
End Function

Private Function ReplaceSpecialChar(ByVal vstrBefore As String) As String
    Dim strAfter As String
    strAfter = vstrBefore
    strAfter = Replace(strAfter, vbCrLf, vbCr)
    strAfter = Replace(strAfter, vbTab, AddBlank(1))

    ReplaceSpecialChar = strAfter
End Function

Private Function ReplaceSpecialCharUndo(ByVal vstrBefore As String) As String
    Dim strAfter As String
    strAfter = vstrBefore
    strAfter = Replace(strAfter, vbCr, vbCrLf)

    ReplaceSpecialCharUndo = strAfter
End Function

Private Function IsSeparate(ByVal vstrChar As String) As Boolean
    If _
        (vstrChar = "") Or _
        (vstrChar = " ") Or _
        (vstrChar = vbTab) Or _
        (vstrChar = vbCr) Then
        IsSeparate = True
    Else
        IsSeparate = False
    End If
End Function

Private Function CQueryKeyword( _
    ByVal vstrBefore As String, _
    ByRef rtypKeyword As Keyword, _
    ByVal vlngPos As Long, _
    ByVal vlngPosNxtCh As Long _
) As String

    Dim strLeft As String
    strLeft = Trim(Left(vstrBefore, vlngPos - 1))
    If strLeft <> "" Then
        If rtypKeyword.SeparateFlag = SEPARATEFLAG_ON Then
            If Right(strLeft, 1) <> vbCr Then
                strLeft = strLeft & vbCr
            End If
        End If
    End If

    Dim strTarget As String
    strTarget = Trim(Mid(vstrBefore, vlngPos, Len(rtypKeyword.Character)))
    strTarget = AddBlank(rtypKeyword.IndentLevel - 1, strTarget)
    strTarget = strTarget & vbCr

    Dim strRight As String
    If Len(vstrBefore) >= vlngPosNxtCh Then
        strRight = Trim(Right(vstrBefore, Len(vstrBefore) - vlngPosNxtCh + 1))
        strRight = Trim(ELeftN(strRight))
        strRight = AddBlank(rtypKeyword.IndentLevel, strRight)
    Else
        strRight = ""
    End If

    CQueryKeyword = strLeft & strTarget & strRight
End Function

Private Function ELeftN(ByVal vstrChar As String) As String
    Dim strRet As String
    strRet = vstrChar

    Do
        Dim strLeft As String
        strLeft = Left(strRet, 1)
        If (strLeft = vbCr) Then
            strRet = Right(strRet, Len(strRet) - 1)
        Else
            Exit Do
        End If
    Loop

    ELeftN = strRet
End Function

Private Function AddBlank( _
    ByVal vintLevel As Integer, _
    Optional ByVal vstrTarget As String = "" _
) As String
    Dim strBlank As String
    strBlank = ""

    Dim i As Integer
    For i = 1 To INDENT_UNIT * vintLevel
        strBlank = strBlank & " "
    Next i

    AddBlank = strBlank & vstrTarget
End Function

Private Sub ShowError(ByVal vobjErr As ErrObject)
    MsgBox CStr(vobjErr.Number) & ":" & vobjErr.Description, vbCritical + vbOKOnly, App.Title
End Sub

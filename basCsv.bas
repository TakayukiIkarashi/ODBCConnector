Attribute VB_Name = "basCsv"
Option Explicit

Public Function GetCsvValue(ByVal vstrCsvLine As String) As Variant
    Dim strRes() As Variant

    Dim lngPos As Long
    lngPos = 1

    Dim i As Integer
    i = 0

    Do
        Dim lngK As Long
        lngK = InStr(lngPos, vstrCsvLine, ",")

        If lngK < 1 Then
            lngK = Len(vstrCsvLine) + 1
        End If

        Dim strVal As String

        Dim strFCh As String
        strFCh = Mid(vstrCsvLine, lngPos, 1)

        If strFCh = """" Then
            Dim lngD As Long
            lngD = InStr(lngPos + 2, vstrCsvLine, """")

            If lngD < 1 Then
                strVal = Mid(vstrCsvLine, lngPos, lngK - lngPos)
                lngPos = lngK + 1
            Else
                strVal = Mid(vstrCsvLine, lngPos + 1, lngD - lngPos - 1)
                lngPos = lngD + 2
            End If
        Else
            strVal = Mid(vstrCsvLine, lngPos, lngK - lngPos)
            lngPos = lngK + 1
        End If

        ReDim Preserve strRes(i)
        strRes(i) = strVal
        i = i + 1

        If lngPos > Len(vstrCsvLine) Then
            Exit Do
        End If
    Loop

    GetCsvValue = strRes()
End Function

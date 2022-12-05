Attribute VB_Name = "modString"
'Split         Split a string into a variant array.
'
'InStrRev      Similar to InStr but searches from end of string.
'
'Replace       To find a particular string and replace it.
'
'Reverse       To reverse a string.
'
Public outtext As String

Public Sub AddText(text As String)
    outtext = outtext & text & vbCrLf
End Sub

Public Function InStrRev(ByVal sIn As String, ByVal _
                                              sFind As String, Optional nStart As Long = 1, _
                         Optional bCompare As VbCompareMethod = vbBinaryCompare) _
                         As Long

Dim nPos As Long

    sIn = Reverse(sIn)
    sFind = Reverse(sFind)

    nPos = InStr(nStart, sIn, sFind, bCompare)
    If nPos = 0 Then
        InStrRev = 0
    Else
        InStrRev = Len(sIn) - nPos - Len(sFind) + 2
    End If
End Function

Public Function Join(Source() As String, _
                     Optional sDelim As String = " ") As String

Dim nC As Long
Dim sOut As String

    For nC = LBound(Source) To UBound(Source) - 1
        sOut = sOut & Source(nC) & sDelim
    Next

    Join = sOut & Source(nC)
End Function

Public Function Replace(ByVal sIn As String, ByVal sFind As _
                                             String, ByVal sReplace As String, Optional nStart As _
                                                                               Long = 1, Optional nCount As Long = -1, _
                        Optional bCompare As VbCompareMethod = vbBinaryCompare) As _
                        String

Dim nC As Long, nPos As Long
Dim nFindLen As Long, nReplaceLen As Long

    nFindLen = Len(sFind)
    nReplaceLen = Len(sReplace)

    If (sFind <> "") And (sFind <> sReplace) Then
        nPos = InStr(nStart, sIn, sFind, bCompare)
        Do While nPos
            nC = nC + 1
            sIn = Left(sIn, nPos - 1) & sReplace & _
                  Mid(sIn, nPos + nFindLen)
            If nCount <> -1 And nC >= nCount Then Exit Do
            nPos = InStr(nPos + nReplaceLen, sIn, sFind, _
                         bCompare)
        Loop
    End If

    Replace = sIn
End Function

Public Function Reverse(ByVal sIn As String) As String
Dim nC As Long
Dim sOut As String

    For nC = Len(sIn) To 1 Step -1
        sOut = sOut & Mid(sIn, nC, 1)
    Next nC

    Reverse = sOut
End Function

Public Function Split(ByVal sIn As String, _
                      Optional sDelim As String = " ", _
                      Optional nLimit As Long = -1, _
                      Optional bCompare As VbCompareMethod = vbBinaryCompare) _
                      As Variant

Dim nC As Long, nPos As Long, nDelimLen As Long
Dim sOut() As String

    If sDelim <> "" Then
        nDelimLen = Len(sDelim)
        nPos = InStr(1, sIn, sDelim, bCompare)
        Do While nPos
            ReDim Preserve sOut(nC)
            sOut(nC) = Left(sIn, nPos - 1)
            sIn = Mid(sIn, nPos + nDelimLen)
            nC = nC + 1
            If nLimit <> -1 And nC >= nLimit Then Exit Do
            nPos = InStr(1, sIn, sDelim, bCompare)
        Loop
    End If

    ReDim Preserve sOut(nC)
    sOut(nC) = sIn

    Split = sOut
End Function


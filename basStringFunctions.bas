Attribute VB_Name = "basStringFunctions"
'String functions for parsing strings.

Option Explicit

Public Function Pc(ByVal vstrSource As String, ByVal vstrDelim As String, ByVal vintBegin As Integer, Optional vintEnd As Integer = 0) As String
'Piece function. Returns a delimited piece of the source string.
    
    Dim intPiece As Integer
    
    If vintEnd = 0 Then
        vintEnd = vintBegin
    End If
    
    Pc = ""
    
    For intPiece = vintBegin To vintEnd
        If intPiece = vintBegin Then
            Pc = MPiece(vstrSource, vstrDelim, intPiece)
        Else
            Pc = Pc & vstrDelim & MPiece(vstrSource, vstrDelim, intPiece)
        End If
    Next

End Function

Public Function NumPc(ByVal vstrSource As String, ByVal vstrDelim As String) As Integer
'Returns the number of delimited pieces in a string.
    
    Dim lngDelimLen As Long
    Dim lngDelimPos As Long
    Dim lngCount As Long
    Dim lngBegPos As Long
    
    lngDelimLen = Len(vstrDelim)
    lngCount = 1
    lngBegPos = 1
    
    Do
        lngDelimPos = InStr(lngBegPos, vstrSource, vstrDelim)
        If lngDelimPos = 0 Then Exit Do
        lngCount = lngCount + 1
        lngBegPos = lngDelimPos + lngDelimLen
    Loop

    NumPc = lngCount

End Function

Private Function MPiece(ByVal vstrSource As String, ByVal vstrDelim As String, ByVal vintPcNum As Integer) As String
    
    Dim lngBegPos As Long
    Dim lngPcLen As Long
    Dim lngDelimLen As Long
    Dim lngDelimCnt As Long
    Dim lngDelimPos As Long
    
    MPiece = ""
    
    If vintPcNum = 1 Then
        lngBegPos = 1
        lngPcLen = InStr(1, vstrSource, vstrDelim) - 1
    Else
        lngDelimLen = Len(vstrDelim)
        lngBegPos = 1
        
        For lngDelimCnt = 1 To vintPcNum - 1
            lngDelimPos = InStr(lngBegPos, vstrSource, vstrDelim)
            If lngDelimPos = 0 Then Exit Function
            lngBegPos = lngDelimPos + lngDelimLen
        Next lngDelimCnt
        
        lngDelimPos = InStr(lngBegPos, vstrSource, vstrDelim)
        
        If lngDelimPos <> 0 Then
            lngPcLen = lngDelimPos - lngBegPos
        Else
            lngPcLen = Len(vstrSource) - lngBegPos + 1
        End If
    End If
    
    If lngPcLen = 0 Then Exit Function
    
    If lngPcLen = -1 Then
        MPiece = vstrSource
        Exit Function
    End If
    
    MPiece = Mid$(vstrSource, lngBegPos, lngPcLen)

End Function

Public Function SetPc(ByVal vstrSource As String, ByVal vstrDelim As String, ByVal vintPcNum As Integer, ByVal vstrReplStr As String) As String
'Set data into a delimited piece of the source string.
    
    Dim intDelimLen As Integer
    Dim intDelimPos As Integer
    Dim intBegPos As Integer
    Dim intCntPc As Integer
    Dim intAddDelim As Integer
    Dim strLeftStr As String
    Dim strRightStr As String
    
    intDelimLen = Len(vstrDelim)
    
    If vintPcNum = 1 Then
        strLeftStr = ""
        intDelimPos = InStr(1, vstrSource, vstrDelim)
        
        If intDelimPos = 0 Then
            strRightStr = ""
        Else
            strRightStr = Right$(vstrSource, (Len(vstrSource) - intDelimPos + 1))
        End If
    Else
        strLeftStr = ""
        intBegPos = 1
        
        For intCntPc = 1 To (vintPcNum - 1)
            intDelimPos = InStr(intBegPos, vstrSource, vstrDelim)
            
            If intDelimPos = 0 Then
                strLeftStr = vstrSource
                For intAddDelim = intCntPc To (vintPcNum - 1)
                    strLeftStr = strLeftStr & vstrDelim
                Next intAddDelim
                Exit For
            Else
                strLeftStr = Left$(vstrSource, (intDelimPos + intDelimLen - 1))
                intBegPos = intDelimPos + intDelimLen
            End If
        Next intCntPc
        
        If intDelimPos = 0 Then
            strRightStr = ""
        Else
            intDelimPos = InStr(intBegPos, vstrSource, vstrDelim)
            
            If intDelimPos = 0 Then
                strRightStr = ""
            Else
                strRightStr = Right$(vstrSource, (Len(vstrSource) - intDelimPos + 1))
            End If
        End If
    End If
    
    SetPc = strLeftStr & vstrReplStr & strRightStr

End Function

Public Function NextPc(ByRef vstrString As String, ByVal vstrDelim As String) As String
'Returns the next piece of a delimited string sequentially, stripping off the returned
'piece each time.
    
    Dim lngDelimPos As Long
    
    lngDelimPos = InStr(1, vstrString, vstrDelim)
    
    If lngDelimPos = 0 Then
        NextPc = vstrString
        vstrString = ""
        Exit Function
    End If
    
    NextPc = Mid$(vstrString, 1, lngDelimPos - 1)
    vstrString = Mid$(vstrString, lngDelimPos + Len(vstrDelim))
    
End Function


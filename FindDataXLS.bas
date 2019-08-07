Function FindDataXLS(ByRef ws As Worksheet, ByRef aTemplates() As String, ByVal rEnd As Integer, ByVal colEnd As Integer, Optional ByVal rInit As Integer = 1, Optional ByVal colInit As Integer = 1, Optional ReturnAddress As Boolean = False, Optional RemoveSymbol As Integer = 0, Optional RemoveTerms As Variant) As String()
            
    Dim ret() As String
    Dim retBase(0) As String
    retBase(0) = ""
    strResults = ""
    
    On Error GoTo ErrChk
    For r = rInit To rEnd
        For col = colInit To colEnd
            strToSplit = ws.Cells(r, col).value
            If (IsError(strToSplit)) Then strToSplit = "#ERRO!"
            If (RemoveSymbol >= 2) Then
                strToSplit = Replace(SymbolTxtConverter(strToSplit), "-", " ")
            End If
            awords = Split(strToSplit, " ")
            For i = LBound(awords) To UBound(awords)
                word = awords(i)
                If Not (IsMissing(RemoveTerms)) Then
                    For Each term In RemoveTerms
                        word = Replace(word, term, "")
                    Next
                End If
                If (RemoveSymbol >= 1) Then word = SymbolTxtRemove(word)
                For T = 0 To UBound(aTemplates)
                    If (LCase(word) Like LCase(aTemplates(T)) And Len(word) = Len(aTemplates(T))) Then
                        If (ReturnAddress) Then
                            strResults = strResults & r & ";" & col & "///"
                        Else
                            strResults = strResults & word & "///"
                        End If
                    End If
                Next
            Next
        Next
    Next
    If (strResults <> "") Then
        ret = Split(Left(strResults, Len(strResults) - 3), "///")
    Else
        ret = retBase
    End If
    FindDataXLS = ret
    
    Exit Function
    
ErrChk:
    MsgBox Err.Description
    Stop
    Resume
    
End Function

<%

    DIM bolCoSort, intNbrKwrds, bolSesKWOnly, strTempKW, bolGenCatOnly, bolGenTxtOnly
    DIM strHldWord, strTypQ, strOper, strLike, bolPhrase, strRP, strLP, bolLastNameFirst, bolMatchWord
    DIM strSrchFlds, bolCatOnly, strCatFlds, strTypeDB

SUB Build_Where_SQL

    DIM strKeyWords, intSub, strReqOper, aryKW(4), intSub2, bolPreOr, strWord, strRParen, bolChkType

    strHldWord = ""
    strWhrSql = ""
    intSub = 0

    If bolCatOnly Then
       intSub = 1
       aryKW(intSub) = REPLACE(SESSION("CAT"),","," OR ")
       strSrchFlds = strCatFlds
       bolMatchWord = true
       SESSION("CAT") = ""
       SESSION("RECCNT") = ""
       If strCatFlds = "" Then
          strCatFlds = "Category"
       End If
    Else
       If not bolSesKWOnly Then
          If SESSION("SRCHKW") <> "" Then
             strKeyWords = SESSION("SRCHKW")
          ElseIf SESSION("emKW") <> "" Then
             strKeyWords = SESSION("emKW")
          ElseIf Request.QueryString("RKW") <> ""  Then
             strKeyWords = Request.QueryString("RKW")
          Else
             strKeyWords = Request.Form("Keywords")
             If SESSION("SRCHKW") <> trim(Request.Form("Keywords")) Then
                Session.Contents.Remove("WHRSQL")
             End If
             SESSION("SRCHKW") = trim(Request.Form("Keywords"))
          End If
       End If

       If strTempKW <> ""  Then
          intSub = 1
          aryKW(intSub) = strTempKW
       End If

       If ucase(strKeyWords) = "*ALL*" Then
          strKeyWords = ""
       End If

       If trim(strKeyWords) <> "" Then
          strKeyWords = RemoveLineBreaks(strKeyWords)
          strKeyWords = TrimSpaces(strKeyWords,1)
          intSub = intSub + 1
          aryKW(intSub) = strKeyWords
       End If
    End If

    strTypeDB = SESSION("PTYP")
    If strTypeDB = "" Then
       strTypeDB = "GEN"
    End If

    aryKW(0) = intSub
    bolDoPhrases = false
    bolChkType = true

    For intSub2 = 1 to aryKW(0)
       intSub = 1
       strWord = GetWords(aryKW(intSub2),intSub,1)
       strHldWord = ""
       Do While strWord <> ""
          If intSub = 1 Then
             strWhrSQL = CheckSrchType
             bolChkType = false
          End If
          If strTypeDB <> "GEN" Then
             If left(strWord,1) = "!"  Then
                'Parms that begin with "!" are searched against all fields, not just Company and Email
                strWord = mid(strWord,2)
             ElseIf (trim(strKeyWords) <> "" and intSub2 = aryKW(0)) or left(aryKW(intSub2),1) = "!" Then
                'If the string begins with "!" all the parms are searched against all fields, not just Company and Email
                'Which means if there is a mix, don't begin the string with a "!" parm
             End If
          End If
          If strHldWord = "" Then
             If intSub > 1 and (ucase(strWord) = "OR" or ucase(strWord) = "AND") Then
                strReqOper = ucase(strWord)
                bolPreOr = true
                strWord = ""
             ElseIf strWhrSql = "" and SESSION("CAT") = "" Then
                strWhrSQL = " WHERE ("
             ElseIf intSub = 1 and intSub2 > 1 Then
                strWhrSQL = strWhrSQL & " AND "
             Else
                If  not bolPreOr Then
                   strReqOper = "AND"
                End If
                strWhrSQL = strWhrSQL & " " & strReqOper & " "
             End If
             If strWord <> "" Then
                If left(strWord,1) = "-" Then
                   strLike = "NOT LIKE"
                   strWord = mid(strWord,2)
                   strOper = "AND"
                Else
                   strLike = "LIKE"
                   strOper = "OR"
                End If
                bolPreOr = false
             End If
          End If
          If strWord <> "" Then
             call Process_Word(strWord)
          End If
          intSub = intSub + 1
          strWord = GetWords(aryKW(intSub2),intSub,1)
       Loop
    Next

    If bolChkType Then
       strWhrSQL = CheckSrchType
    End If

    If strWhrSQL <> "" Then
       strWhrSQL = strWhrSQL & ")"
    End If

    If not bolCatOnly and SESSION("CAT") <> "" Then
       If strWhrSQL <> "" Then
          strWhrSQL = left(SESSION("CAT"),len(SESSION("CAT"))-1) & strWhrSQL
       Else
          strWhrSQL = SESSION("CAT")
       End If
    End If

    If SESSION("DEBUG") = "Y" Then
       Response.Write "Where SQL " &  strWhrSql & "<br>" & vbCrLf
    End If

    If bolCatOnly Then
       SESSION("CAT") = strWhrSQL
       strWhrSQL = ""
       bolCatOnly = false
    End If

End SUB

SUB Process_Word(DATA)

    DIM strWord, intLoc

    strWord = DATA

    intLoc = instr(strWord,CHR(10))
    If intLoc > 0 Then
       strWord = left(strWord,intLoc-1)
    End If
    intLoc = instr(strWord,CHR(13))
    If intLoc > 0 Then
       strWord = left(strWord,intLoc-1)
    End If
    If left(strWord,1) = "(" Then
       strLP = "("
       strWord = mid(strWord,2)
    ElseIf right(strWord,1) = ")" Then
       strRP = ")"
       strWord = left(strWord,len(strWord)-1)
    End If
    If strHldWord = "" Then
       bolPhrase = false
       strTypQ = ""
    End If
    If (left(strWord,1) = "'" or left(strWord,1) = CHR(34)) and strTypQ = "" Then
        strTypQ = left(strWord,1)
        bolPhrase = true
        If right(strWord,1) = strTypQ Then
           strWord = mid(left(strWord,len(strWord)-1),2)
        Else
           strHldWord = mid(strWord,2)
           strWord = ""
        End If
    ElseIf strHldWord <> "" and right(strWord,1) = strTypQ Then
        strWord = strHldWord & " " & left(strWord,len(strWord)-1)
        strHldWord = ""
    End If
    If strHldWord = "" Then
       strWhrSQL = strWhrSQL & strLP & WhrString(strWord,strLike,strOper) & strRP
       intNbrKwrds = intNbrKwrds + 1
       strLP = ""
       strRP = ""
    ElseIf strWord <> "" Then
       strHldWord = strHldWord & " " & strWord
    End If

END SUB

FUNCTION WhrString(CurWord,CurLike,CurOper)

    DIM strSep, intCnt, strHld, strEqLogic, strPS, strLikeLogic, strFldName
    DIM strHldWrdSep, bolHldPhraseOption, bolLocMatchWord

    strPS = "%"
    strLikeLogic = CurLike
    bolLocMatchWord = bolMatchWord

    If CurLike = "LIKE" Then
       strEqLogic = "="
    Else
      strEqLogic = "<>"
    End If

    strSep = "'"
    If instr(CurWord,strSep) > 0 Then
       For intCnt = 1 to len(CurWord)
           If mid(CurWord,intCnt,1) = "'" Then
              strHld = strHld & "''"
           Else
              strHld = strHld & mid(CurWord,intCnt,1)
           End If
       Next
       CurWord = strHld
    end If

    If CurWord = "." Then
       CurWord = ""
       strLikeLogic = "="
       strPS = ""
    End If
    If left(CurWord,1) = "*" or right(CurWord,1) = "*" Then
       strPS = ""
       If left(CurWord,1) = "*" Then
          CurWord = "%" & mid(CurWord,2)
       End If
       If right(CurWord,1) = "*" Then
          CurWord = left(CurWord,len(CurWord)-1) & "%"
       End If
    End If

    WhrString = "("

    If strTypeDB = "SE" or strTypeDB = "SG" Then
       WhrString = WhrString & "[Security Groups] " & strLikeLogic & " " & strSep & strPS & ":" & CurWord & ":"  & strPS & strSep & " "
    ElseIf strTypeDB = "GEN" Then
       If not bolGenCatOnly Then
          If bolLocMatchWord then
             WhrString = WhrString & "[Text] " & strEqLogic & " " & strSep & CurWord & strSep
          Else
             WhrString = WhrString & "[Text] " & strLikeLogic & " " & strSep & strPS & CurWord & strPS & strSep
          End If
          If  not bolGenTxtOnly Then
             WhrString = WhrString & " "& CurOper & " "
          End If
       End If
       If  not bolGenTxtOnly Then
          WhrString = WhrString & "[Category] " & strEqLogic & " " & strSep & CurWord & strSep
       End If
    ElseIf strsrchFlds <> "" Then

       strHldWrdSep = strWrdSep
       bolHldPhraseOption = bolDoPhrases
       bolDoPhrases = false
       strWrdSep = ";"
       intCnt = 1
       strFldName = GetWords(strSrchFlds,intCnt,1)
       do while strFldName <> ""
          If left(strFldName,1) = "#" Then
             strFldName = mid(strFldName,2)
             If bolMatchWord Then
                bolLocMatchWord = false
             Else
                bolLocMatchWord = true
             End If
          Else
             bolLocMatchWord = bolMatchWord
          End If
          If WhrString <> "(" Then
             WhrString = WhrString & " "& CurOper & " "
          End If
          If bolLocMatchWord then
             WhrString = WhrString & "[" & strFldName & "] " & strEqLogic & " " & strSep & CurWord & strSep
          Else
             WhrString = WhrString & "[" & strFldName & "] " & strLikeLogic & " " & strSep & strPS & CurWord & strPS & strSep
          End If
          intCnt = intCnt + intRetWords
          strFldName = GetWords(strSrchFlds,intCnt,1)
       Loop
    Else
       WhrString = ""
    End If
    If WhrString <> "" Then
       WhrString = WhrString & ")"
    End If
    strWrdSep = strHldWrdSep
    bolDoPhrases = bolHldPhraseOption
END FUNCTION

FUNCTION CheckSrchType

   If trim(SESSION("SRCHTYPE")) <> "" and not bolCatOnly Then
      If SESSION("CAT") = "" Then
         CheckSrchType = " WHERE ("
      Else
         CheckSrchType = CheckSrchType & " AND "
      End If
      If SESSION("SRCHTYPE") = "A" Then
         CheckSrchType = CheckSrchType & "[Inactive] = false"
      ElseIf SESSION("SRCHTYPE") = "I" Then
         CheckSrchType = CheckSrchType & "[Inactive] = true"
      End If
   End If

END FUNCTION
%>


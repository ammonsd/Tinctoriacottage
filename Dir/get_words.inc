
<%
     DIM bolDoPhrases, intRetWords, bolQuoteIsPhrase, strWrdSep

FUNCTION GetWords(DATA,START_WORD,NBR_WORDS)

    DIM intCur, intLoc, strCurData, intNbrSel, strW, strWords, bolIsPhrase, strTemp, intNbrWords

    If strWrdSep = "" or isnull(strWrdSep) Then
       strWrdSep = " "
    Else
       strWrdSep = left(strWrdSep,1)
    End If

    If DATA = "" or not isnumeric(START_WORD) or not isnumeric(NBR_WORDS) Then
       GetWords = DATA
       EXIT FUNCTION
    End IF

    intCur = 0
    intNbrSel = 0
    strCurData = trim(DATA)
    intNbrWords = NBR_WORDS
    strWords = ""
    intRetWords = 0

    If bolQuoteIsPhrase Then
       bolDoPhrases = false
    End If

    If bolDoPhrases and (instr(strCurData,",") > 0 or instr(strCurData,strWrdSep) > 0) Then
       strCurData = strCurData & ","
       strTemp = ""
       Do while instr(strCurData,",") > 0
          intLoc = instr(strCurData,",")
          strW = trim(left(strCurData,intLoc-1))
          strCurData = trim(mid(strCurData,intLoc + 1))
          If left(strW,1) <> CHR(34) and right(strW,1) <> CHR(34) and instr(strW,strWrdSep) > 0 Then
             strW = CHR(34) & strW & CHR(34)
          End If
          If strTemp = "" Then
             strTemp = strW
          Else
             strTemp = strTemp & strWrdSep & strW
          End If
       Loop
       strCurData = trim(strTemp)
    End If

    Do while strCurData <> ""
       intCur = intCur + 1
       intLoc = instr(strCurData,strWrdSep)
       If intLoc = 0 Then
          intLoc = len(strCurData) + 1
          intNbrWords = 0
       End If
       If intCur => START_WORD Then
          strW = left(strCurData,intLoc-1)
          If not bolDoPhrases and bolQuoteIsPhrase and left(strW,1) = CHR(34) Then
             bolDoPhrases = true
          End If
          If bolDoPhrases Then
             If left(strW,1) = CHR(34) and not bolIsPhrase Then
                strW = mid(strW,2)
                bolIsPhrase = true
             End If
             If right(strW,1) = CHR(34) and bolIsPhrase Then
                strW = left(strW,len(strW)-1)
                bolIsPhrase = false
             End If
          End If
          If strWords = "" Then
             strWords = strW
          Else
             strWords = strWords & strWrdSep & strW
          End If
          intRetWords = intRetWords + 1
          If not bolIsPhrase Then
             intNbrSel = intNbrSel + 1
          End If
          If intNbrSel => intNbrWords Then
             EXIT DO
          End If
       End If
       strCurData = trim(mid(strCurData,intLoc+1))
    Loop
    GetWords = strWords
END FUNCTION

FUNCTION TotalWords(DATA)

    DIM strWord

    TotalWords = 0
    strWord = " "

    do while strWord <> ""
       TotalWords = TotalWords + 1
       strWord = GetWords(DATA,TotalWords,1)
    Loop
    TotalWords = TotalWords - 1

END FUNCTION
%>

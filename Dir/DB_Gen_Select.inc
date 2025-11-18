
<!--#INCLUDE FILE="../dir/array_sort.inc" -->

<%
    DIM aryGenText(5000,11), aryGenText2(5000,8), aryGenMemo(5000,4), aryGenDteNum(5000,5), bolMore, strPageDir
    DIM intMaxRecs, intFRec, strTotLine, intPgNbr, intPgs, intLID, strSrchSeq, strSrchArea, bolNoSrch, bolDoSearch

SUB Get_Gen_Text_Recs(TYP,CAT)

    DIM intSub, strHld, intCurRec, bolFirst, bolRecOK, strSQL, bolNewRecCnt, strWord, strOper

    bolDoPhrases = false
    Erase aryGenText

    If intMaxRecs = "" or not isnumeric(intMaxRecs) or intMaxRecs > 5000 Then
       intMaxRecs = 5000
    End If

    If strPageDir = "-" Then
       strPageDir = "B"
       intPgNbr = intPgNbr + 1
    End If

    If strPageDir = "" or strPageDir = "T" Then
       intPgNbr = 0
       intPgs = ""
       intLID = 0
    End If

    If strPageDir = "B" Then
       If intPgNbr > 1 Then
          intPgNbr = intPgNbr - 1
          intFRec = int(GetWords(intPgs,intPgNbr,1))
          intPgs = GetWords(intPgs,1,intPgNbr)
       End If
    Else
       intPgNbr = intPgNbr + 1
       intFRec = int(intLID) + 1
    End If

    Call Database_Setup

    If TYP <> "" Then
       If isNumeric(TYP) Then
          bolNoSrch = true
       End If
    End If

    If strSrchSeq <> "" Then
       bolDoSearch = true
    End IF

    If SESSION("SRCHSQL") <> "" and not bolNoSrch and not bolDoSearch Then
       strHld =  SESSION("SRCHSQL")
    ElseIf (strWhrSql <> "" or bolDoSearch) and not bolNoSrch Then
       strHld = SearchIDs(TYP,CAT)
       SESSION("SRCHSQL") = strHld
       bolNewRecCnt = true
    End If
    If SESSION("SRCHSQL") = "" or bolNoSrch Then
       If not bolNoSrch Then
          bolNewRecCnt = true
       End If
       strHld = " WHERE [FKEY] = 0" 'Top Level Entries Only
       If TYP <> "" Then
          If isNumeric(TYP) Then
             strHld = strHld & " AND [ID] = " & TYP
          Else
             strHld = strHld & " AND [TYPE] = '" & TYP & "'"
          End If
       End If
       If CAT <> "" Then
          strOper = "IN"
          CAT = REPLACE(CAT,","," ")
          intSub = 1
          strWord = GetWords(CAT,intSub,1)
          do while strWord <> ""
             If left(strWord,1) = "-" Then
                strWord = mid(strWord,2)
                strOper = "NOT IN"
             End If
             If intSub = 1 Then
                strSQL = "'" & strWord & "'"
             Else
                strSql = strSql & ",'" & strWord & "'"
             End If
             intSub = intSub + 1
             strWord = GetWords(CAT,intSub,1)
          loop
          strHld = strHld & " AND [CATEGORY] " & strOper & " (" & strSQL & ")"
       End If
       If strWhrSql <> "" Then
          strHld = strHld & " AND " & mid(strWhrSql,8)
       End If
    End If

    Set objRecordSet = Server.CreateObject("ADODB.Recordset")
    objRecordSet.ActiveConnection = strDbConn

    If SESSION("RECCNT") = "" or bolNewRecCnt or Request.Form("SEARCH") = "Y" Then
       strSQL = "SELECT Count(*) AS RecCnt FROM " & APPLICATION("TEXTTBL") & strHld
       objRecordSet.Open strSQL
       SESSION("RECCNT") = objRecordSet("RecCnt")
       objRecordSet.Close
    End If

    If SESSION("RECCNT") > 0 Then
       strTotLine = intFRec & "-"
       If intMaxRecs = 5000 or (intFRec+intMaxRecs-1) > SESSION("RECCNT") Then
          strTotLine = strTotLine & SESSION("RECCNT")
       Else
          strTotLine = strTotLine & intFRec+intMaxRecs-1
       End If
       strTotLine = strTotLine & " of " & SESSION("RECCNT")
    End If

    intSub = 0
    intCurRec = 0

    If SESSION("SRCHSQL") <> "" and SESSION("RECCNT") = 0 Then
    Else
       objRecordSet.Source = "SELECT * FROM " & APPLICATION("TEXTTBL") & strHld & " ORDER BY [TYPE], [SEQUENCE], [TEXT]"
       objRecordSet.LockType = adLockReadOnly
       objRecordSet.Open

       If Not objRecordSet.EOF Then
          objRecordSet.MoveFirst
          bolMore = false
          bolFirst = true
          Do While Not objRecordSet.EOF
             bolRecOK = true
             intCurRec = intCurRec + 1
             If intCurRec < intFRec Then
                bolRecOK = false
             ElseIf SESSION("SRCHTYPE") = "A" and objRecordSet.Fields("INACTIVE") Then
                bolRecOK = false
             ElseIf SESSION("SRCHTYPE") = "I" and not objRecordSet.Fields("INACTIVE") Then
                bolRecOK = false
             End If
             If bolRecOK Then
                If bolFirst Then
                   bolFirst = false
                   If strPageDir <> "B" Then
                      If intPgs = "" Then
                         intPgs = intCurRec
                      Else
                         intPgs = intPgs & " " & intCurRec
                      End If
                   End If
                End If
                intSub = intSub + 1
                aryGenText(intSub,1) = objRecordSet.Fields("ID")
                aryGenText(intSub,2) = trim(objRecordSet.Fields("TYPE"))
                aryGenText(intSub,3) = trim(objRecordSet.Fields("CATEGORY"))
                aryGenText(intSub,4) = objRecordSet.Fields("SEQUENCE")
                aryGenText(intSub,5) = trim(objRecordSet.Fields("TEXT"))
                aryGenText(intSub,6) = objRecordSet.Fields("INACTIVE")
                aryGenText(intSub,7) = objRecordSet.Fields("FKEY")
                aryGenText(intSub,8) = objRecordSet.Fields("ADDED")
                aryGenText(intSub,9) = objRecordSet.Fields("UPDATED")
                aryGenText(intSub,10) = objRecordSet.Fields("ADDED USER ID")
                aryGenText(intSub,11) = objRecordSet.Fields("UPDATE USER ID")
                If intSub = intMaxRecs Then
                   objRecordSet.MoveNext
                   If Not objRecordSet.EOF Then
                      bolMore = true
                   End If
                   EXIT DO
                End If
             End If
             objRecordSet.MoveNext
          Loop
          objRecordset.Close
       End If
    End If

    Set objRecordSet = Nothing

    aryGenText(0,0) = intSub
    intLID = intCurRec
    strWhrSql = ""

END SUB

SUB Get_Gen_Text2_Recs(ID)

    DIM intSub, strHld

    Call Database_Setup
    Erase aryGenText2

    Set objRecordSet = Server.CreateObject("ADODB.Recordset")
    objRecordSet.ActiveConnection = strDbConn

    If ID <> "" and isNumeric(ID) Then
       strHld = " WHERE [FKEY] = " & ID & " OR [FKEY2] = " & ID
    Else
       strHld = " WHERE [CATEGORY] = '" & ID & "'"
    End If

    objRecordSet.Source = "SELECT * FROM " & APPLICATION("TEXTTBL") & strHld & " ORDER BY [SEQUENCE],[TEXT]"

    objRecordSet.LockType = adLockReadOnly
    objRecordSet.Open

    intSub = 0

    If Not objRecordSet.EOF Then
       objRecordSet.MoveFirst
       Do While Not objRecordSet.EOF
          intSub = intSub + 1
          aryGenText2(intSub,1) = objRecordSet.Fields("ID")
          aryGenText2(intSub,2) = trim(objRecordSet.Fields("TYPE"))
          aryGenText2(intSub,3) = objRecordSet.Fields("SEQUENCE")
          aryGenText2(intSub,4) = trim(objRecordSet.Fields("TEXT"))
          aryGenText2(intSub,5) = objRecordSet.Fields("INACTIVE")
          aryGenText2(intSub,6) = objRecordSet.Fields("FKEY")
          aryGenText2(intSub,7) = objRecordSet.Fields("FKEY2")
          aryGenText2(intSub,8) = trim(objRecordSet.Fields("CATEGORY"))
          objRecordSet.MoveNext
       Loop
    End If

    aryGenText2(0,0) = intSub

    objRecordset.Close
    Set objRecordSet = Nothing

END SUB

SUB Get_Gen_Nbr_Date_Recs(ID)

    DIM intSub

    Call Database_Setup
    Erase aryGenDteNum

    Set objRecordSet = Server.CreateObject("ADODB.Recordset")
    objRecordSet.ActiveConnection = strDbConn

    If ID = "" or not isNumeric(ID) Then
       ID = 0
    End If

    objRecordSet.Source = "SELECT * FROM " & APPLICATION("NBRDTETBL") & " WHERE [FKEY] = " & ID & " ORDER BY [SEQUENCE]"

    objRecordSet.LockType = adLockReadOnly
    objRecordSet.Open

    intSub = 0
    If Not objRecordSet.EOF Then
       objRecordSet.MoveFirst
       Do While Not objRecordSet.EOF
          intSub = intSub + 1
          aryGenDteNum(intSub,1) = objRecordSet.Fields("ID")
          aryGenDteNum(intSub,2) = objRecordSet.Fields("SEQUENCE")
          aryGenDteNum(intSub,3) = objRecordSet.Fields("NUMBER")
          aryGenDteNum(intSub,4) = objRecordSet.Fields("DATE")
          aryGenDteNum(intSub,5) = objRecordSet.Fields("FKEY")
          objRecordSet.MoveNext
       Loop
    End If

    aryGenDteNum(0,0) = intSub

    objRecordset.Close
    Set objRecordSet = Nothing

END SUB

SUB Get_Gen_Memo_Recs(ID)

    DIM intSub, strHld

    Call Database_Setup
    Erase aryGenMemo

    Set objRecordSet = Server.CreateObject("ADODB.Recordset")
    objRecordSet.ActiveConnection = strDbConn

    If ID = "" or not isNumeric(ID) Then
       ID = 0
    End If
    strHld = " WHERE [FKEY] = " & ID
    objRecordSet.Source = "SELECT * FROM " & APPLICATION("MEMOTBL") & strHld & " ORDER BY [SEQUENCE]"

    objRecordSet.LockType = adLockReadOnly
    objRecordSet.Open

    intSub = 0
    If Not objRecordSet.EOF Then
       objRecordSet.MoveFirst
       Do While Not objRecordSet.EOF
          intSub = intSub + 1
          aryGenMemo(intSub,1) = objRecordSet.Fields("ID")
          aryGenMemo(intSub,2) = objRecordSet.Fields("SEQUENCE")
          aryGenMemo(intSub,3) = trim(objRecordSet.Fields("TEXT"))
          aryGenMemo(intSub,4) = objRecordSet.Fields("FKEY")
          objRecordSet.MoveNext
       Loop
    End If

    aryGenMemo(0,0) = intSub

    objRecordset.Close
    Set objRecordSet = Nothing

END SUB

FUNCTION SearchIDs(TYP,CAT)

    DIM strHld, aryCnt(50000), intSub, intCnt, intFRng, intLRng, strIDs, intID, strSQL, intPrID, strCatSQL

    Call Database_Setup

    Set objRecordSet = Server.CreateObject("ADODB.Recordset")
    objRecordSet.ActiveConnection = strDbConn
    strIDS = ""
    intSub = -1

    If CAT <> "" Then
       strCatSQL = " AND [CATEGORY] = '" & CAT & "'"
    End If

    If instr("C1",strSrchArea) > 0 or strSrchArea = "" Then
       strHld = "SELECT DISTINCT([ID]),[FKEY] FROM " & APPLICATION("TEXTTBL")
       strHld = strHld & " Where ([TYPE] = '" & TYP & "' AND [FKEY] = 0"
       If strCatSQL <> "" Then
          strHld = strHld & strCatSQL
       End If
       strHld = strHld & ")"
       If strWhrSql <> "" Then
          strHld = strHld & " AND " & mid(strWhrSql,8)
       End If

       If SESSION("DEBUG") = "Y" Then
          Response.Write "Text#1 Srch Sql = " & strHld & "<br><br>" & vbCrLf
       End If

       objRecordSet.Source = strHld

       objRecordSet.LockType = adLockReadOnly
       objRecordSet.Open

       If Not objRecordSet.EOF Then
          objRecordSet.MoveFirst
          Do While Not objRecordSet.EOF
             intSub = intSub + 1
             aryCnt(intSub) = objRecordSet.Fields("ID")
             If strIDs = "" Then
                strIDS = aryCnt(intSub)
             Else
                strIDS = strIDS & "," & aryCnt(intSub)
             End If
             objRecordSet.MoveNext
          Loop
       End If
       objRecordset.Close
    End If

    If instr("C2",strSrchArea) > 0 or strSrchArea = "" Then
       strHld = "SELECT DISTINCT([ID]),[FKEY] FROM " & APPLICATION("TEXTTBL")
       strHld = strHld & " Where ([TYPE] = '" & TYP & "' AND [FKEY] > 0"
       If strSrchSeq <> "" Then
          strHld = strHld & " AND [SEQUENCE] IN (" & strSrchSeq & ")"
       End If
       If strIDs <> "" Then
          If instr(strWhrSql," NOT LIKE ") = 0 Then
             strHld = strHld & " AND [FKEY] NOT IN (" & strIDs & ")"
          Else
             strHld = strHld & " AND [FKEY] IN (" & strIDs & ")"
             strIDs = ""
             Erase aryCnt
             intSub = -1
          End If
       End If
       If strCatSQL <> "" Then
          strHld = strHld & strCatSQL
       End If
       strHld = strHld & ")"
       If strWhrSql <> "" Then
          strHld = strHld & " AND " & mid(strWhrSql,8)
       End If


       If SESSION("DEBUG") = "Y" Then
          Response.Write "Text#2 Srch Sql = " & strHld & "<br><br>" & vbCrLf
       End If

       objRecordSet.Source = strHld

       objRecordSet.LockType = adLockReadOnly
       objRecordSet.Open

       If Not objRecordSet.EOF Then
          objRecordSet.MoveFirst
          Do While Not objRecordSet.EOF
             intSub = intSub + 1
             aryCnt(intSub) = objRecordSet.Fields("FKEY")
             If strIDs = "" Then
                strIDS = aryCnt(intSub)
             Else
                strIDS = strIDS & "," & aryCnt(intSub)
             End If
             objRecordSet.MoveNext
          Loop
       End If

       objRecordset.Close
    End If

    If strSrchArea = "M" or strSrchArea = "" Then
       strHld = "SELECT DISTINCT([ID]),[FKEY] FROM " & APPLICATION("MEMOTBL")
       strHld = strHld & " Where ([TYPE] = '" & TYP & "'"
       If strIDs <> "" Then
          If instr(strWhrSql," NOT LIKE ") = 0 Then
             strHld = strHld & " AND [FKEY] NOT IN (" & strIDs & ")"
          Else
             strHld = strHld & " AND [FKEY] IN (" & strIDs & ")"
             strIDs = ""
             Erase aryCnt
             intSub = -1
          End If
       End If
       If strCatSQL <> "" Then
          strHld = strHld & strCatSQL
       End If
       strHld = strHld & ")"
       If strWhrSql <> "" Then
          strHld = strHld & " AND " & mid(strWhrSql,8)
       End If

       If SESSION("DEBUG") = "Y" Then
          Response.Write "Memo Srch Sql = " & strHld & "<br><br>" & vbCrLf
       End If

       objRecordSet.Source = strHld

       objRecordSet.LockType = adLockReadOnly
       objRecordSet.Open

       If Not objRecordSet.EOF Then
          objRecordSet.MoveFirst
          Do While Not objRecordSet.EOF
             intSub = intSub + 1
             aryCnt(intSub) = objRecordSet.Fields("FKEY")
             If strIDs = "" Then
                strIDS = aryCnt(intSub)
             Else
                strIDS = strIDS & "," & aryCnt(intSub)
             End If
             objRecordSet.MoveNext
          Loop
       End If

       objRecordset.Close
    End If

    intFRng = 0
    intLRng = 0
    strIDS = ""
    strSQL = ""

    arySorted = SortArray(aryCnt,0)
    Erase aryCnt

    If intNbrEntries > 0 Then
       arySorted(intNbrEntries) = 999999999999999999999999
       intNbrEntries = intNbrEntries + 1
    End If

    For intSub = 0 to intNbrEntries - 1
       intID = int(arySorted(intSub))
       If intID <> intPrID Then
          intPrID = intID
          If intFRng = 0 Then
             intFRng = intID
          ElseIf (intLRng = 0 and intID = intFRng+1) or (intLRng > 0 and intID = intLRng+1) Then
             intLRng = intID
          ElseIf intLRng=0 Then
             If strIDs = "" Then
                strIDS = intFRng
             Else
                strIDS = strIDS & "," & intFRng
             End If
             intFRng = intID
          Else
             IF strSQL = "" Then
                strSQL = " ("
             Else
                strSQL = strSQL & " OR "
             End If
             strSQL = strSQL & "[ID] BETWEEN " & intFRng & " AND " & intLRng
             intFRng = intID
             intLRng = 0
          End If
       End If
    Next

    Set objRecordSet = Nothing
    Erase arySorted

    If strIDs <> "" Then
       If strSQL <> "" Then
          strSQL = strSQL & " OR"
       Else
          strSQL = " ("
       End If
       strSQL = strSQL & " [ID] IN (" & strIDS & ")"
    End IF

    If strSQL = "" Then
       strSQL = " ([ID] IN (0)"
    End If

    SearchIDs = " WHERE " & strSQL & ")"

    If SESSION("DEBUG") = "Y" Then
       Response.Write "Search Where SQL = " & SearchIDs & "<br><br>" & vbCrLf
    End If

END FUNCTION
%>


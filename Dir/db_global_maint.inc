
<!--#INCLUDE FILE="../dir/write_errors.inc" -->

<%

    DIM aryRecDtls(20,3), aryRecData(5000,20), aryDbDtls(4), bolMore, strPageDir
    DIM intMaxRecs, intFRec, strTotLine, intPgNbr, intPgs, intLID, strTypeSel

FUNCTION AddGlobalRecord()

   DIM strTblName, intSub, strFldName, strData, strRecType

   strTblName = aryDbDtls(1)

   Set objRecordSet = Server.CreateObject("ADODB.Recordset")
   objRecordSet.Open strTblName, strDbConn, adOpenStatic, adLockPessimistic, adCmdTable
   on error resume next
   objRecordSet.AddNew

   For intSub = 1 to UBOUND(aryRecDtls)
      If trim(aryRecDtls(intSub,1)) = "" Then
         EXIT FOR
      End If
      strFldName = aryRecDtls(intSub,1)
      strData = aryRecDtls(intSub,2)
      strRecType = aryRecDtls(intSub,3)
      If strData <> CHR(254) Then
         If strRecType = "TS" Then
            objRecordSet.Fields(strFldName) = NOW()
         ElseIf strRecType = "I" Then
            If strData = "" or not isNumeric(strData) Then
               strDaat = "0"
            End If
            objRecordSet.Fields(strFldName) = int(strData)
         Else
            objRecordSet.Fields(strFldName) = strData
         End If
      End If
   Next
   objRecordSet.Update

   objRecordSet.MoveLast
   aryDbDtls(0) = objRecordSet.Fields(aryRecDtls(1,1))

   objRecordSet.Close
   Set objRecordSet = Nothing

   If Err.Number <> 0 and Err.Number <> 53 Then
      AddGlobalRecord = false
      call Write_Errors
   Else
      AddGlobalRecord = true
   End If

END FUNCTION

FUNCTION UpdateGlobalRecord()

   DIM strTblName, intSub, strFldName, strData, intRecID

   intRecID = int(aryDbDtls(0))
   strTblName = aryDbDtls(1)

   Set objRecordSet = Server.CreateObject("ADODB.Recordset")
   objRecordSet.Open strTblName, strDbConn, adOpenStatic, adLockPessimistic, adCmdTable

   objRecordSet.MoveFirst
   objRecordSet.Find ("ID = " & intRecID)

   UpdateGlobalRecord = true

   If objRecordSet.EOF Then
      UpdateGlobalRecord = false
   Else
      For intSub = 1 to UBOUND(aryRecDtls)
         If trim(aryRecDtls(intSub,1)) = "" Then
            EXIT FOR
         End If
         strFldName = aryRecDtls(intSub,1)
         strData = aryRecDtls(intSub,2)
         If strData = "[TS]" Then
            objRecordSet.Fields(strFldName) = NOW()
         ElseIf strData <> CHR(254) Then
            objRecordSet.Fields(strFldName) = strData
         End If
      Next
      objRecordSet.Update
   End If

   objRecordSet.Close
   Set objRecordSet = Nothing

   If Err.Number <> 0 and Err.Number <> 53 Then
      UpdateGlobalRecord = false
      call Write_Errors
   End If

END FUNCTION

FUNCTION DeleteGlobalRecord

   DIM strTblName, intRecID, strFldName

   intRecID = int(aryDbDtls(0))
   strTblName = aryDbDtls(1)
   strFldName = aryRecDtls(1,1)

   Set objRecordSet = Server.CreateObject("ADODB.Recordset")
   objRecordSet.ActiveConnection = strDbConn
   objRecordSet.Source = "DELETE FROM " & strTblName & " WHERE [" & strFldName & "] = " & intRecID
   objRecordSet.LockType = adLockReadOnly
   objRecordSet.Open
   Set objRecordSet = Nothing

   If Err.Number <> 0 and Err.Number <> 53 and Err.Number <> 53 Then
      DeleteGlobalRecord = false
      call Write_Errors
   Else
      DeleteGlobalRecord = true
   End If

END FUNCTION

SUB Get_Global_Records

    DIM intSub, intSub2, strHld, intCurRec, bolFirst, strSQL, strWord, strOper
    DIM strTblName, strSortFlds, intRecID, strRecID, strFldName

    strRecID = aryDbDtls(0)
    strTblName = aryDbDtls(1)
    strSortFlds = aryDbDtls(2)

    bolDoPhrases = false

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

    If strRecID <> "" and strRecID <> "0" Then
       intRecID = int(strRecID)
       strFldName = aryRecDtls(1,1)
       strWhrSQL = " WHERE [" & strFldName & "] = " & intRecID
    End If

    Set objRecordSet = Server.CreateObject("ADODB.Recordset")
    objRecordSet.ActiveConnection = strDbConn

    If SESSION("RECCNT") = "" or Request.Form("SEARCH") = "Y" Then
       strSQL = "SELECT Count(*) AS RecCnt FROM " & strTblName & strWhrSQL
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

    If strSortFlds <> "" Then
       strWrdSep = ","
       bolDoPhrases = false
       intSub = 1
       strWord=GetWords(strSortFlds,intSub,1)
       strHld = ""
       do while strWord <> ""
          If isNumeric(strWord) Then
             intSub2 = int(strWord)
             strFldName = aryRecDtls(intSub2,1)
             If strHld = "" Then
                strHld = " ORDER BY [" & strFldName & "]"
             Else
                strHld = strHld & ",[" & strFldName & "]"
             End If
          End If
          intSub = intSub + 1
          strWord=GetWords(strSortFlds,intSub,1)
       Loop
       strWhrSQL = strWhrSQL & strHld
    End If

    intSub = 0
    intCurRec = 0

    objRecordSet.Source = "SELECT * FROM " & strTblName & strWhrSQL
    objRecordSet.LockType = adLockReadOnly
    objRecordSet.Open

    If Not objRecordSet.EOF Then
       objRecordSet.MoveFirst
       bolMore = false
       bolFirst = true
       Do While Not objRecordSet.EOF
          intCurRec = intCurRec + 1
          If intCurRec >= intFRec Then
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
             For intSub2 = 1 to UBOUND(aryRecDtls)
                If trim(aryRecDtls(intSub2,1)) = "" Then
                   EXIT FOR
                End If
                strFldName = aryRecDtls(intSub2,1)
                aryRecData(intSub,intSub2) = objRecordSet.Fields(strFldName)
             Next
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
    End If

    objRecordset.Close
    Set objRecordSet = Nothing

    aryRecData(0,0) = intSub
    intLID = intCurRec
    strWhrSql = ""

END SUB

%>


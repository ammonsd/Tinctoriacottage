
<!--#INCLUDE FILE="../dir/write_errors.inc" -->

<%

   DIM aryGenRec(500,3), intLastSeq, intFKey

FUNCTION AddGenTextRec(TYP,CAT,SEQ,TEXT,STATUS,FKEY,FKEY2)

   DIM bolStatus

   AddGenTextRec = 0

   Set objRecordSet = Server.CreateObject("ADODB.Recordset")
   objRecordSet.Open APPLICATION("TEXTTBL"), strDbConn, adOpenStatic, adLockPessimistic, adCmdTable
   on error resume next
   objRecordSet.AddNew

   If SEQ = "" or not isNumeric(SEQ) Then
      SEQ = 0
   End If

   If STATUS = true or STATUS = "true" or STATUS = "0" Then
      bolStatus = true
   Else
      bolStatus = false
   End If

   objRecordSet.Fields("TYPE") = TYP
   objRecordSet.Fields("CATEGORY") = CAT
   objRecordSet.Fields("SEQUENCE") = SEQ
   objRecordSet.Fields("TEXT") = TEXT
   objRecordSet.Fields("INACTIVE") = bolStatus
   If FKEY <> "" and isNumeric(FKEY) Then
      objRecordSet.Fields("FKEY") = int(FKEY)
   End If
   If FKEY2 <> "" and isNumeric(FKEY2) Then
      objRecordSet.Fields("FKEY2") = int(FKEY2)
   End If
   objRecordSet.Fields("Added") = NOW()
   objRecordSet.Fields("Added User ID") = SESSION("USERID")

   objRecordSet.Update

   objRecordSet.MoveLast
   AddGenTextRec = objRecordSet.Fields("ID")

   objRecordSet.Close

END FUNCTION

FUNCTION AddGenMemoRec(TYP,CAT,SEQ,TEXT,FKEY)

   Set objRecordSet = Server.CreateObject("ADODB.Recordset")
   objRecordSet.Open APPLICATION("MEMOTBL"), strDbConn, adOpenStatic, adLockPessimistic, adCmdTable
   on error resume next
   objRecordSet.AddNew

   objRecordSet.Fields("TYPE") = TYP
   objRecordSet.Fields("CATEGORY") = CAT
   objRecordSet.Fields("SEQUENCE") = SEQ
   objRecordSet.Fields("TEXT") = TEXT
   objRecordSet.Fields("FKEY") = int(FKEY)

   objRecordSet.Fields("Added") = NOW()
   objRecordSet.Fields("Added User ID") = SESSION("USERID")

   objRecordSet.Update

   objRecordSet.MoveLast
   AddGenMemoRec = objRecordSet.Fields("ID")

   objRecordSet.Close

END FUNCTION

FUNCTION AddGenNbrDteRec(TYP,SEQ,NBR,DTE,FKEY)

   Set objRecordSet = Server.CreateObject("ADODB.Recordset")
   objRecordSet.Open APPLICATION("NBRDTETBL"), strDbConn, adOpenStatic, adLockPessimistic, adCmdTable
   on error resume next
   objRecordSet.AddNew

   objRecordSet.Fields("TYPE") = TYP
   objRecordSet.Fields("SEQUENCE") = SEQ
   If NBR <> "" and isNumeric(NBR) Then
      objRecordSet.Fields("NUMBER") = NBR
   End If
   IF DTE <> "" and isDate(DTE) Then
      objRecordSet.Fields("DATE") = DTE
   End If
   objRecordSet.Fields("FKEY") = int(FKEY)

   objRecordSet.Fields("Added") = NOW()
   objRecordSet.Fields("Added User ID") = SESSION("USERID")

   objRecordSet.Update

   objRecordSet.MoveLast
   AddGenNbrDteRec = objRecordSet.Fields("ID")

   objRecordSet.Close

END FUNCTION

FUNCTION UpdGenTextRec(ID,CAT,SEQ,TEXT,STATUS)

   Set objRecordSet = Server.CreateObject("ADODB.Recordset")
   objRecordSet.Open APPLICATION("TEXTTBL"), strDbConn, adOpenStatic, adLockPessimistic, adCmdTable

   If SEQ = "" or not isNumeric(SEQ) Then
      SEQ = 0
   End If

   If objRecordSet.EOF Then
       UpdGenTextRec = false
   Else
     objRecordSet.MoveFirst
     objRecordSet.Find ("ID = " & ID)
     If objRecordSet.EOF Then
        UpdGenTextRec = false
     Else
        UpdGenTextRec = true
        If CAT <> "" Then
           objRecordSet.Fields("CATEGORY") = CAT
        End If
        If SEQ > 0 Then
           objRecordSet.Fields("SEQUENCE") = SEQ
        End If
        If TEXT <> "" Then
           objRecordSet.Fields("TEXT") = TEXT
        End If
        If STATUS <> "" Then
           objRecordSet.Fields("INACTIVE") = STATUS
        End If
        objRecordSet.Fields("Updated") = NOW()
        objRecordSet.Fields("Update User ID") = SESSION("USERID")
     End If
   End If

   objRecordSet.Update
   objRecordSet.Close
   Set objRecordSet = Nothing

   Call UpdGenMemoCat(ID,CAT)

END FUNCTION

FUNCTION UpdGenMemoRec(ID,TEXT)

   Set objRecordSet = Server.CreateObject("ADODB.Recordset")
   objRecordSet.Open APPLICATION("MEMOTBL"), strDbConn, adOpenStatic, adLockPessimistic, adCmdTable

   If objRecordSet.EOF Then
       UpdGenMemoRec = false
   Else
     objRecordSet.MoveFirst
     objRecordSet.Find ("ID = " & ID)
     If objRecordSet.EOF Then
        UpdGenMemoRec = false
     Else
        UpdGenMemoRec = true
        objRecordSet.Fields("TEXT") = TEXT
        objRecordSet.Fields("Updated") = NOW()
        objRecordSet.Fields("Update User ID") = SESSION("USERID")
     End If
   End If

   objRecordSet.Update
   objRecordSet.Close
   Set objRecordSet = Nothing

END FUNCTION

FUNCTION UpdGenMemoCat(FKEY,CAT)

   DIM strSQL

   Set objRecordSet = Server.CreateObject("ADODB.Recordset")
   objRecordSet.LockType = adLockPessimistic
   objRecordSet.CursorLocation = adUseClient
   objRecordSet.CursorType = adOpenStatic
   strSQL = "UPDATE " & APPLICATION("MEMOTBL") & " SET [CATEGORY] = '" & CAT & "' WHERE [FKEY] = " & FKEY
   objRecordSet.Open strSQL, strDbConn
   Set objRecordSet = Nothing

END FUNCTION

FUNCTION UpdGenNbrDteRec(ID,NBR,DTE)

   Set objRecordSet = Server.CreateObject("ADODB.Recordset")
   objRecordSet.Open APPLICATION("NBRDTETBL"), strDbConn, adOpenStatic, adLockPessimistic, adCmdTable

   If objRecordSet.EOF Then
       UpdGenNbrDteRec = false
   Else
     objRecordSet.MoveFirst
     objRecordSet.Find ("ID = " & ID)
     If objRecordSet.EOF Then
        UpdGenNbrDteRec = false
     Else
        UpdGenNbrDteRec = true
        If NBR <> "" and isNumeric(NBR) Then
           objRecordSet.Fields("NUMBER") = NBR
        End If
        IF DTE <> "" and isDate(DTE) Then
           objRecordSet.Fields("DATE") = DTE
        End If
        objRecordSet.Fields("Updated") = NOW()
        objRecordSet.Fields("Update User ID") = SESSION("USERID")
     End If
   End If

   objRecordSet.Update
   objRecordSet.Close
   Set objRecordSet = Nothing

END FUNCTION

FUNCTION CascadeDelGenTextRec(ID)

   DIM intSub

   Call Get_Gen_Text2_Recs(ID)
   Call Get_Gen_Nbr_Date_Recs(ID)
   Call Get_Gen_Memo_Recs(ID)

   For intSub = 1 to aryGenText2(0,0)
      Call DelGenTextRec(aryGenText2(intSub,1))
   Next
   For intSub = 1 to aryGenMemo(0,0)
      Call DelGenMemoRec(aryGenMemo(intSub,1))
   Next
   For intSub = 1 to aryGenDteNum(0,0)
      Call DelGenNbrDteRec(aryGenDteNum(intSub,1))
   Next
   Call DelGenTextRec(ID)

END FUNCTION

FUNCTION DelGenTextRec(ID)

   Set objRecordSet = Server.CreateObject("ADODB.Recordset")
   objRecordSet.ActiveConnection = strDbConn
   objRecordSet.Source = "DELETE FROM " & APPLICATION("TEXTTBL") & " WHERE [ID] = " & ID
   objRecordSet.LockType = adLockReadOnly
   objRecordSet.Open
   Set objRecordSet = Nothing
   DelGenTextRec = true

END FUNCTION

FUNCTION DelGenMemoRec(ID)

   Set objRecordSet = Server.CreateObject("ADODB.Recordset")
   objRecordSet.ActiveConnection = strDbConn
   objRecordSet.Source = "DELETE FROM " & APPLICATION("MEMOTBL") & " WHERE [ID] = " & ID
   objRecordSet.LockType = adLockReadOnly
   objRecordSet.Open
   Set objRecordSet = Nothing
   DelGenMemoRec = true

END FUNCTION

FUNCTION DelGenNbrDteRec(ID)

   Set objRecordSet = Server.CreateObject("ADODB.Recordset")
   objRecordSet.ActiveConnection = strDbConn
   objRecordSet.Source = "DELETE FROM " & APPLICATION("NBRDTETBL") & " WHERE [ID] = " & ID
   objRecordSet.LockType = adLockReadOnly
   objRecordSet.Open
   Set objRecordSet = Nothing
   DelGenNbrDteRec = true

END FUNCTION

FUNCTION AddGenTextRecords(TYP,CAT,STATUS)

   DIM strData, strData2, intSeq, strTypRec, intSub

   AddGenTextRecords = true

   intFKey = AddGenTextRec(TYP,CAT,0,aryGenRec(0,1),STATUS,"","")

   If Err.Number <> 0 and Err.Number <> 53 Then
      AddGenTextRecords = false
      call Write_Errors
      EXIT FUNCTION
   End If

   If intFKey > 0 Then
      strData = ""
      For intSub = 1 to intLastSeq
         intSeq = aryGenRec(intSub,0)
         strData = aryGenRec(intSub,1)
         strTypRec = aryGenRec(intSub,2)
         strData2 = aryGenRec(intSub,3)
         If strData <> "" or (strData2 <> "" and strTypRec = "N") Then
            If strTypRec = "T" Then
               Call AddGenTextRec(TYP,CAT,intSeq,strData,STATUS,intFKey,"")
            ElseIf strTypRec = "N" Then
               Call AddGenNbrDteRec(TYP,intSeq,strData,strData2,intFKey)
            ElseIf strTypRec = "M" Then
               Call AddGenMemoRec(TYP,CAT,intSeq,strData,intFkey)
            End If

            If Err.Number <> 0 and Err.Number <> 53 Then
               AddGenTextRecords = false
               call Write_Errors
               EXIT FUNCTION
            End If

         End If
      Next
   End If

END FUNCTION

FUNCTION UpdateGenTextRecords(ID,TYP,CAT,STATUS)

   DIM intSub, intSub2, intSeq, strData, strData2, bolRecFnd(500), intRecID, strTypRec

   UpdateGenTextRecords = true

   Call UpdGenTextRec(ID,CAT,"",aryGenRec(0,1),STATUS)

   If Err.Number <> 0 and Err.Number <> 53 Then
      UpdateGenTextRecords = false
      call Write_Errors
      EXIT FUNCTION
   End If

   Call Get_Gen_Text2_Recs(ID)
   Call Get_Gen_Nbr_Date_Recs(ID)
   Call Get_Gen_Memo_Recs(ID)

   For intSub = 1 to aryGenText2(0,0)
      intRecID = aryGenText2(intSub,1)
      intSeq = aryGenText2(intSub,3)
      For intSub2 = 1 to intLastSeq
         If intSeq = aryGenRec(intSub2,0) and aryGenRec(intSub2,2) = "T" Then
            strData = aryGenRec(intSub2,1)
            bolRecFnd(intSub2) = true
            If strData = "" Then
               Call DelGenTextRec(intRecID)
            Else
               Call UpdGenTextRec(intRecID,CAT,"",strData,STATUS)
            End If
            If Err.Number <> 0 and Err.Number <> 53 Then
               UpdateGenTextRecords = false
               call Write_Errors
               EXIT FUNCTION
            Else
               EXIT FOR
            End If
         End If
      Next
   Next

   For intSub = 1 to aryGenDteNum(0,0)
      intRecID = aryGenDteNum(intSub,1)
      intSeq = aryGenDteNum(intSub,2)
      For intSub2 = 1 to intLastSeq
         If intSeq = aryGenRec(intSub2,0) and aryGenRec(intSub2,2) = "N" Then
            strData = aryGenRec(intSub2,1)
            strData2 = aryGenRec(intSub2,3)
            bolRecFnd(intSub2) = true
            If strData = "" and strData2 = "" Then
               Call DelGenNbrDteRec(intRecID)
            Else
               Call UpdGenNbrDteRec(intRecID,strData,strData2)
            End If
            If Err.Number <> 0 and Err.Number <> 53 Then
               UpdateGenTextRecords = false
               call Write_Errors
               EXIT FUNCTION
            Else
               EXIT FOR
            End If
         End If
      Next
   Next

   For intSub = 1 to aryGenMemo(0,0)
      intRecID = aryGenMemo(intSub,1)
      intSeq = aryGenMemo(intSub,2)
      For intSub2 = 1 to intLastSeq
         If intSeq = aryGenRec(intSub2,0) and aryGenRec(intSub2,2) = "M" Then
            strData = aryGenRec(intSub2,1)
            bolRecFnd(intSub2) = true
            If strData = "" Then
               Calll elGenMemoRec(intRecID)
            Else
               Call UpdGenMemoRec(intRecID,strData)
            End If
            If Err.Number <> 0 and Err.Number <> 53 Then
               UpdateGenTextRecords = false
               call Write_Errors
               EXIT FUNCTION
            Else
               EXIT FOR
            End If
         End If
      Next
   Next

   For intSub = 1 to intLastSeq
      intSeq = aryGenRec(intSub,0)
      strData = aryGenRec(intSub,1)
      strTypRec = aryGenRec(intSub,2)
      strData2 = aryGenRec(intSub,3)
      If (strData <> "" or strData2 <> "" and strTypRec = "N") and not bolRecFnd(intSub) Then
         If strTypRec = "T" Then
            Call AddGenTextRec(TYP,CAT,intSeq,strData,STATUS,ID,"")
         ElseIf strTypRec = "N" Then
            Call AddGenNbrDteRec(TYP,intSeq,strData,strData2,ID)
         ElseIf strTypRec = "M" Then
            Call AddGenMemoRec(TYP,CAT,intSeq,strData,ID)
         End If
         If Err.Number <> 0 and Err.Number <> 53 and Err.Number <> 53 Then
            UpdateGenTextRecords = false
            call Write_Errors
            EXIT FUNCTION
         End If
      End If
   Next

END FUNCTION
%>


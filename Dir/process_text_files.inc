<%
   DIM aryTxtRecs(1000), bolReturnComments, bolAppendFile, bolDelFile

SUB Read_Text_File(TxtFn,DIR)

   DIM strTextFn, objTextFn, intCntr, intSub, strRec

   aryTxtRecs(0) = 0
   bolDoPhrases = false

   If DIR = "" Then
      strTextFn = ReplChars(Request.ServerVariables("Script_Name"),"/"," ")
      strTextFn = "\" & GetWords(strTextFn,1,1) & "\"
      strTextFn = Server.Mappath(strTextFn) & "\" & TxtFn
   Else
      on error resume next
      strTextFn = Server.Mappath("\" & DIR & "\") & "\" & TxtFn
   End If

   If err.Number <> 0 Then
      EXIT SUB
   End If

   Set objTextFn = CreateObject("Scripting.FileSystemObject")
   on error resume next
   Set objTextFn = objTextFn.OpenTextFile(strTextFn)

   intSub = 0

   if err.Number = 0 then
      Do While Not objTextFn.AtEndOfStream and intSub < 1001
         strRec = objTextFn.ReadLine
         If left(strRec,1) = ":" and not bolReturnComments Then
           'Skip Comment Record
         Else
            intSub = intSub + 1
            If intSub > UBound(aryTxtRecs) Then
               intSub = intSub -1
               EXIT DO
            End If
            aryTxtRecs(intSub) = strRec
         End If
      Loop
      objTextFn.close
      set objTextFn = nothing
   End if

   aryTxtRecs(0) = intSub

END SUB

FUNCTION CheckFileExist(FN)

   DIM objTextFn, strFn

   strFn = FN
   If mid(strFn,2,1) <> ":" Then
      strFn = FullFileName(strFn)
   End If
   Set objTextFn = CreateObject("Scripting.FileSystemObject")
   on error resume next
   Set objTextFn = objTextFn.OpenTextFile(strFn)

   CheckFileExist = err.Number

   SET ObjTextFn = nothing
END FUNCTION

FUNCTION FullFileName(FN)

   DIM strSvrName, intLoc

   strSvrName = ucase(Request.ServerVariables("SERVER_NAME"))
   FullFileName = FN

   intLoc = instr(ucase(FullFileName),strSvrName)
   If intLoc > 0 Then
      FullFileName = mid(FullFileName,intLoc + len(strSvrName))
   End If

   FullFileName = Server.Mappath("\") & "\" & FullFileName

END FUNCTION

SUB Write_Text_File(TxtFn,DIR)

   DIM strTextFn, objTextFn, ObjFSO, intCntr, intSub, bolCreate, intTypeUpd

   bolDoPhrases = false

   If DIR = "" Then
      strTextFn = ReplChars(Request.ServerVariables("Script_Name"),"/"," ")
      strTextFn = "\" & GetWords(strTextFn,1,1) & "\"
      strTextFn = Server.Mappath(strTextFn) & "\" & TxtFn
   Else
      strTextFn = Server.Mappath("\" & DIR & "\") & "\" & TxtFn
   End If

   If bolAppendFile Then
      intTypeUpd = 8
   Else
      intTypeUpd = 2
   End If
   bolCreate = true

   Set objFSO = CreateObject("Scripting.FileSystemObject")
   on error resume next
   Set objTextFn = objFSO.OpenTextFile(strTextFn,intTypeUpd,bolCreate)

   intSub = 0

   if err.Number > 0 then
      Call Write_Errors
   Else
      bolDelFile = true
      For intSub = 1 to aryTxtRecs(0)
         If trim(aryTxtRecs(intSub)) <> "" Then
            bolDelFile = false
            EXIT FOR
         End If
      Next
      IF bolDelFile Then
         objTextFn.close
         Set objTextFn = objFSO.GetFile(strTextFn)
         objTextFn.Delete
      Else
         For intSub = 1 to aryTxtRecs(0)
            objTextFn.WriteLine aryTxtRecs(intSub)
         Next
         objTextFn.close
      End If
      set objFSO=nothing
      set objTextFn=nothing
   End if

   if err.Number > 0 then
      Call Write_Errors
   End If

END SUB
%>


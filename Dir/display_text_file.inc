<!--#INCLUDE FILE="../dir/update_text_variables.inc" -->

<%
   DIM strMsgVar, strDisSubj

SUB Display_Text_File(TxtFn)

   DIM strTextFn, objTextFn, intCntr, intSub, strText, bolSkip

   bolDoPhrases = false
   strTextFn = ReplChars(Request.ServerVariables("Script_Name"),"/"," ")
   strTextFn = "\" & GetWords(strTextFn,1,1) & "\"

   strTextFn = Server.Mappath(strTextFn) & "\" & TxtFn

   Set objTextFn = CreateObject("Scripting.FileSystemObject")
   on error resume next
   Set objTextFn = objTextFn.OpenTextFile(strTextFn)

   If strMsgVar <> "" Then
      SESSION(strMsgVar) = ""
   End If

   if err.Number = 0 then
      Response.Write "<html><head><title></title></head><body>"
      Do While Not objTextFn.AtEndOfStream
         strText = CheckTextVariables(objTextFn.ReadLine)
         If instr(strText,"[SKIP]") > 0 Then
            If bolSkip Then
               bolSkip = false
            Else
               bolSkip = true
            End If
         ElseIf not bolSkip Then
            If left(ucase(strText),3) = "SB:" Then
               strDisSubj = mid(strText,4)
            Else
               If strMsgVar <> "" Then
                  If SESSION(strMsgVar) = "" Then
                     SESSION(strMsgVar) = strText
                  Else
                     SESSION(strMsgVar) = SESSION(strMsgVar) & " " & strText
                  End If
               End If
               Response.Write strText
            End If
         End If
      Loop
      Response.Write "</body></html>"
      objTextFn.close
      set objTextFn=nothing
   End if

END SUB%>


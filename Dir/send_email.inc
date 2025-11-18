
<!--#INCLUDE FILE="../dir/update_text_variables.inc" -->

<%

   DIM bolPreMsg, strGoPgm

FUNCTION SendJMail(SMTPServer,ToAddr,CCAddr,BCCAddr, FromName, FromAddr, Subj, Msg, FileAtt, FileAttAsIs, RetRec)

   DIM objJMail, strTemp, strEM, strName, intLoc, intCID, strEmHTML1, strEmHTML2, intSub, strSigHTML, strSigEM
   DIM bolNoSig, strMsg, bolNoHTML, strUcMsg, strTemp2, intLoc2, strDate

   strMsg = Msg
   strUcMsg = ucase(strMsg)

   If left(strUcMsg,7) = "<DEFER>" and instr(strUcMsg,"</DEFER>") > 0 Then
      intLoc = instr(strUcMsg,"</DEFER>")
      strDate = left(strMsg,intLoc-1)
      strDate = mid(strDate,8)
      strMsg = mid(strMsg,intLoc+8)
      strUcMsg = ucase(strMsg)
      If not isDate(strDate) Then
         strDate = ""
      End If
   End If

   If instr(strUcMsg,"FROM:")>0 and instr(strUcMsg,"TO:")>0 and instr(strUcMsg,"SUBJECT:")>0 Then
      'Assume this is a reply to a message and the signature is not required.
      bolNoSig = true
   End If

   If instr(strUcMsg,"<HTML") > 0 Then
      bolNoHTML = true
      bolNoSig = true
   End If

   If left(FromAddr,1) = "-" Then
      bolNoSig = true
      FromAddr = mid(FromAddr,2)
   End If

   Call Database_Setup

   If SMTPServer = "" Then
      SMTPServer = strSMTPServer
   End If

   set objJMail = Server.CreateObject("JMail.Message")

   objJMail.Silent = true

   If strDate <> "" Then
      objJMail.DeferredDelivery = CDate(strDate)
   End If

   objJMail.From = FromAddr
   objJMail.FromName = FromName
   objJMail.Subject = Subj
   If left(strUcMsg,6) = "<TEXT>" Then
      strMsg = mid(strMsg,7)
      If bolPreMsg Then
        Call Preview_Message(ReplLineBreaks(strMsg))
        EXIT FUNCTION
      Else
         objJMail.Body = strMsg
      End If
   Else
      If Not bolNoHTML Then
         call Read_Text_File(SESSION("USERID")&"_EmHTML.CFG",strFileLoc)
         bolDoPhrases = false
         For intSub = 1 to aryTxtRecs(0)
            If left(aryTxtRecs(intSub),1) = "<" Then
               If strEmHTML1 = "" Then
                  strEmHTML1 = aryTxtRecs(intSub)
               End If
            ElseIf not bolNoSig Then
               strSigEM = lcase(GetWords(aryTxtRecs(intSub),1,1))
               If left(strSigEm,1) = "-" and mid(strSigEm,2) = FromAddr Then
                  strSigEm = ""
                  EXIT FOR
               End If
               If FromAddr = strSigEM or strSigEM = "all" Then
                  strSigHTML = strSigHTML & GetWords(aryTxtRecs(intSub),2,99)
               End If
            End If
         Next
      End If
      intLoc = instr(strEmHTML1,"</")
      If intLoc = 0 Then
         strEmHTML1 = ""
      Else
         strEmHTML2 = mid(strEmHTML1,intLoc)
         strEmHTML1 = left(strEmHTML1,intLoc-1)
      End If


      If strEmHTML1 <> "" and instr(lcase(strEmHTML1),"<body") = 0 Then
         strEmHTML1 = "<body>" & strEmHTML1
      End If

      If Not bolNoHTML and instr(strUcMsg,"<HTML") = 0 Then
         If instr(strUcMsg,"<BR>") = 0 Then
            strMsg = ReplLineBreaks(strMsg)
         End If
         If instr(strUcMsg,"</") = 0 Then
            strMsg = ReplMultiBlanks(strMsg)
         End If
      End If

      IF bolNoHTML Then
         strEmHTML1 = CheckTextVariables(strMsg)
      Else
         strEmHTML1 = CheckTextVariables("<html>" & strEmHTML1 & strMsg & strEmHTML2 & strSigHTML & "</body></html>")
      End If

      If bolPreMsg Then
        Call Preview_Message(strEmHTML1)
        EXIT FUNCTION
      End If
      objJMail.HTMLBody = strEmHTML1

   End If
   objJMail.ReturnReceipt = RetRec

   If FileAtt <> "" Then
      If left(ucase(FileAtt),5) = "HTTP:" Then
         strTemp = FileAtt
         strTemp2 = FileAttAsIs
         do while strTemp <> ""
            intLoc = instr(strTemp,";")
            intLoc2 = instr(strTemp2,";")
            If intLoc=0 Then
               intLoc = len(strTemp)+1
               intLoc2 = len(strTemp2)+1
            End If
            intCID = objJMail.AddUrlAttachment(left(strTemp,intLoc-1),left(strTemp2,intLoc2-1))
            strTemp = mid(strTemp,intLoc+1)
            strTemp2 = mid(strTemp2,intLoc2+1)
         Loop
      Else
         FileAtt = FullFileName(ReplChars(FileAtt,"/","\"))
         objJMail.AddAttachment(FileAtt)
      End IF
   End If

   strTemp = ReplChars(ToAddr,";",",")
   If left(strTemp,1) = "," Then
      strTemp = mid(strTemp,2)
   End If
   Do
      intLoc = instr(strTemp,",")
      If intLoc = 0 Then
         intLoc = len(strTemp) + 1
      End If
      strEM = left(strTemp,intLoc-1)
      strTemp = mid(strTemp,intLoc+1)
      intLoc = instr(strEM,"<")
      If intLoc > 0 Then
         strName = left(strEM,intLoc-1)
         strEM = trim(ReplChars(mid(strEM,intLoc+1),">"," "))
      Else
         strName = ""
      End If
      objJMail.AddRecipient strEM,strName
   Loop Until strTemp = ""

   strTemp = ReplChars(CCAddr,";",",")
   If left(strTemp,1) = "," Then
      strTemp = mid(strTemp,2)
   End If
   Do while strTemp <> ""
      intLoc = instr(strTemp,",")
      If intLoc = 0 Then
         intLoc = len(strTemp) + 1
      End If
      strEM = left(strTemp,intLoc-1)
      strTemp = mid(strTemp,intLoc+1)
      intLoc = instr(strEM,"<")
      If intLoc > 0 Then
         strName = left(strEM,intLoc-1)
         strEM = trim(ReplChars(mid(strEM,intLoc+1),">"," "))
      Else
         strName = ""
      End If
      objJMail.AddRecipientCC strEM,strName
   Loop

   strTemp = ReplChars(BCCAddr,";",",")
   If left(strTemp,1) = "," Then
      strTemp = mid(strTemp,2)
   End If
   Do while strTemp <> ""
      intLoc = instr(strTemp,",")
      If intLoc = 0 Then
         intLoc = len(strTemp) + 1
      End If
      strEM = left(strTemp,intLoc-1)
      strTemp = mid(strTemp,intLoc+1)
      intLoc = instr(strEM,"<")
      If intLoc > 0 Then
         strName = left(strEM,intLoc-1)
         strEM = trim(ReplChars(mid(strEM,intLoc+1),">"," "))
      Else
         strName = ""
      End If
      objJMail.AddRecipientBCC strEM,strName
   Loop

   on error resume next
   objJMail.Send(SMTPServer)
   SendJMail = err.Number
   set objJMail = Nothing

END FUNCTION

SUB Preview_Message(MSG)

   Response.Write MSG & vbCrLf
   Response.Write "<div align='center'><br><br>" & vbCrLf
   If strGoPgm = "" Then
      Response.Write "<INPUT TYPE='button' VALUE='BACK' onClick='history.go(-1)'>" & vbCrLf
   Else
      Response.Write "<input type=" & CHR(34) & "button" & CHR(34) & " value=" & CHR(34) & "BACK" & CHR(34) & " onClick=" & CHR(34) & "document.location.href = '" & strGoPgm & "'" & CHR(34) & ">" & vbCrLf
   End if
   Response.Write "</div>" & vbCrLf
   Response.End

END SUB

FUNCTION SendEmail(SMTPServer,ToAddr, FromName, FromAddr, Subj, Msg)

   DIM objSMTP, strTemp, intLoc, strToAddr

   If SMTPServer = "" Then
      SMTPServer = strSMTPServer
   End If

   strTemp = ToAddr
   set objSMTP = Server.CreateObject("Bamboo.SMTP")

   Do
      intLoc = instr(strTemp,";")
      If intLoc = 0 Then
         strToAddr = strTemp
         strTemp=""
      Else
         strToAddr = left(strTemp,intLoc-1)
         strTemp = mid(strTemp,intLoc+1)
      End If

      intLoc = instr(strToAddr,":")
      If intLoc > 0 Then
         strToAddr = mid(strToAddr,intLoc+1)
      End If

      intLoc = instr(strToAddr,"<")
      If intLoc > 0 Then
         strToAddr = mid(strToAddr,intLoc+1)
      End If

      intLoc = instr(strToAddr,">")
      If intLoc > 0 Then
         strToAddr = left(strToAddr,intLoc-1)
      End If

      objSMTP.Server = SMTPServer
      objSMTP.Rcpt = strToAddr
      objSMTP.FromName = FromName
      objSMTP.From = FromAddr
      objSMTP.Subject = Subj
      objSMTP.Message = Msg
      on error resume next
      objSMTP.Send
      SendEmail = err.Number

   Loop Until strTemp = ""

   set objSMTP = Nothing

END FUNCTION
%>

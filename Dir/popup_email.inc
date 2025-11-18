<%

    DIM strToAddr, strCCAddr, strBCCAddr, strCCName, bolNoBCC, bolNoDisplay

FUNCTION SpecialEmOptions(REQ)

    DIM strHld, intSub, intSub1, intSub2, strEm, strEM2, strName, strHld2, strEMN, aryAddrs(3), bolEMFound, intFactor

    strSpecSecGrp = "WebMail"
    bolDisplayMsg = false
    Call Verify_Group_Access("N/R","",9)
    If REQ = "BLDEM" and not bolSecError Then
      SpecialEmOptions = "<input type=" & CHR(34) & "button" & CHR(34) & " value=" & CHR(34) & "Open Web Mail" & CHR(34) & " onClick=" & CHR(34) & "document.location.href = 'email.asp'" & CHR(34) & ">"
      aryAddrs(1) = strToAddr
      If strCCAddr <> "" Then
         aryAddrs(2) = strCCAddr
      End If
      If strBCcAddr <> "" Then
         aryAddrs(3) = strBCcAddr
      End If
      If SESSION("EMN") <> "" Then
         bolDoPhrases = false
         bolQuoteIsPhrase = false
         strWrdSep = ""
         strEMN = ReplChars(SESSION("EMN"),","," ")
         For intSub = 1 to 3
            strHld = ReplChars(aryAddrs(intSub),";"," ")
            intSub1 = 1
            strEm = GetWords(strHld,intSub1,1)
            strHld2 = ""
            intFactor = 2
            Do While strEm <> ""
               intSub2 = 1
               strEM2 = GetWords(strEMN,intSub2,1)
               Do While strEM2 <> ""
                  If strEM2 = strEm Then
                     intSub2 = intSub2 + 1
                     strName = GetWords(strEMN,intSub2,1)
                     If strHld2 = "" Then
                        strHld2 = strName & " {" & strEm & "}"
                     Else
                        strHld2 = strHld2 & ", " & strName & " {" & strEm & "}"
                     End If
                     strEm = ""
                     intFactor = 1
                     EXIT DO
                  End If
                  intSub2 = intSub2 + intFactor
                  strEM2 = GetWords(strEMN,intSub2,1)
               Loop
               If strEm <> "" Then
                  If strHld2 = "" Then
                     strHld2 = strEM
                  Else
                     strHld2 = strHld2 & ", " & strEm
                  End If
               End If
               intSub1 = intSub1 + 1
               strEm = GetWords(strHld,intSub1,1)
            Loop
            If strHld2 <> "" Then
               aryAddrs(intSub) = strHld2
            End If
         Next
      End If
      strHld = ""
      For intSub = 1 to 3
         If aryAddrs(intSub) <> "" Then
            If intSub = 1 Then
                SESSION("TO") = ReplMultiChars(aryAddrs(intSub))
            ElseIf intSub = 2 Then
                SESSION("CC") = ReplMultiChars(aryAddrs(intSub))
            ElseIf intSub = 3 Then
                SESSION("BCC") = ReplMultiChars(aryAddrs(intSub))
            End If
         End If
      Next
    End If

END FUNCTION

Function ReplMultiChars(DATA)
   ReplMultiChars = ReplChars(DATA,";",",")
   ReplMultiChars = ReplChars(ReplMultiChars,"<","{")
   ReplMultiChars = ReplChars(ReplMultiChars,">","}")
   ReplMultiChars = ReplChars(ReplMultiChars,"~"," ")
END FUNCTION

'Pop Up a Preformatted eMail Message Window
SUB Popup_eMail(TOADR,CCADR,BCCADR,SUBJ,MSG)

  DIM strHld

  strHld = Format_MailTo(TOADR,CCADR,BCCADR,SUBJ,MSG)

  If not bolNoDisplay Then
     Response.Write "<div align='center'><br>" & vbCrLf
     Response.Write "<form>" & vbCrLf
     Response.Write "<input type=" & CHR(34) & "button" & CHR(34) & " value=" & CHR(34) & "Open Mail" & CHR(34) & " onClick=" & CHR(34) & "document.location.href = '" & strHld & "'" & CHR(34) & ">" & vbCrLf
  End If
  strHld = SpecialEmOptions("BLDEM")
  If not bolNoDisplay Then
     If strHld <> "" Then
        Response.Write strHld & vbCrLf
     End If
     If SESSION("MENUFN") <> "" Then
        Response.Write "<input type=" & CHR(34) & "button" & CHR(34) & " value=" & CHR(34) & "Return To Menu" & CHR(34) & " onClick=" & CHR(34) & "document.location.href = '" & SESSION("MENUFN") & "'" & CHR(34) & ">" & vbCrLf
     End If
     Response.Write "</form></div>" & vbCrLf
  End If
END SUB

FUNCTION Format_MailTo(TOADR,CCADR,BCCADR,SUBJ,MSG)

  MSG  = ReplChars(MSG, "&", "+")

  If trim(SUBJ) = "" Then
     SUBJ = " "
  Else
     SUBJ = ReplChars(SUBJ, "&", "+")
  End If

  Format_MailTo = "mailto:" & TOADR & "?subject=" & SUBJ

  If trim(CCADR) <> "" Then
      Format_MailTo = Format_MailTo & "&cc=" & CCADR
  End If
  IF trim(BCCADR) <> "" Then
      Format_MailTo = Format_MailTo & "&bcc=" & BCCADR
  End If
  IF trim(MSG) <> "" Then
     Format_MailTo = Format_MailTo & "&body=" & ReplParmChars(MSG)
  End If

END FUNCTION

SUB Popup_eMail_SS(TOADR,CCADR,BCCADR,SUBJ,MSG)
'Server Side Pop Up Preformatted email Message Window

  DIM strHld

  If trim(SUBJ) = "" or isnull(SUBJ) Then
     SUBJ = "No Subject"
  Else
     SUBJ = ReplChars(SUBJ, "&", "+")
  End If

  MSG  = ReplChars(MSG, "&", "+")
  MSG  = ReplChars(MSG, CHR(9), " ")

  strHld = "mailto:" & TOADR & "?subject=" & SUBJ

  IF trim(CCADR) <> "" Then
     strHld = strHld & "&cc=" & CCADR
  End If
  IF trim(BCCADR) <> "" Then
      strHld = strHld & "&bcc=" & BCCADR
  End If
  strHld = strHld & "&body=" & ReplParmChars(MSG)
  If len(strHld) > 150 Then
     strHld = left(strHld,150)
  End If
  Response.ReDirect(strHld)

END SUB

%>


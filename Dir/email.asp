<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
Server.ScriptTimeout = 99999
%>
<!--#INCLUDE FILE="../dir/maint_webpage.inc" -->
<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/send_email.inc" -->
<!--#INCLUDE FILE="../dir/cookie_maint.inc" -->
<!--#INCLUDE FILE="../dir/form_data.inc" -->
<!--#INCLUDE FILE="../dir/get_folder_details.inc" -->
<!--#INCLUDE FILE="../dir/get_data.inc" -->
<!--#INCLUDE FILE="../dir/email_msg_variables.inc" -->
<!--#INCLUDE FILE="../dir/check_dupe_email.inc" -->
<!--#INCLUDE FILE="../dir/process_setup.inc" -->
<!--#INCLUDE FILE="../dir/select_entry.inc" -->
<%

   DIM intErrNbr, strRecFnd, strFrName, strFrEmail, strToEmail, strBCSender, strEmailMsg, strTemp, strSubj, intLoc
   DIM aryEM(3), intSub, strCCEmail, strBCCEmail, strFileAtt, strHdr, strEM, strPgm, strMsgFn, bolRR, strHld, strMsg
   DIM bolAttFnOption, strSenderEM, intCntr, X, strFSpec, strFileAttAsIs, strFN, strTitle, intRows, bolIndivMailing
   DIM aryFileAtt(20), aryFileAttAsIs(20), bolBldShortCut, bolGetAttFile, strAFSpec, strTextMsg, bolBcMailList
   DIM strDupeChkEmail

   strGoPgm = "email.asp"

   Call System_Setup(strGoPgm)

   If FormData("IM") <> "" Then
      SESSION("IM") = FormData("IM")
   End If

   If SESSION("ML") <> "" Then
      bolBcMailList = true
   End If

   If SESSION("IM") <> "" Then
      bolIndivMailing = true
   End If

   If Request.Form("Reset Form") <> "" Then
      bolNoFormData = true
      For Each X in Request.Form
         Session.Contents.Remove(X)
      Next
      Session.Contents.Remove("EMTP")
   ElseIf Request.Form("FormData") = "Y" Then
      bolFormDataOnly = true
   End if

   strSubj = FormData("Subj")
   strToEmail = FormData("TO")
   strCCEmail = FormData("CC")
   strBCCEmail = FormData("BCC")
   strFrName = FormData("SN")
   strFrEmail = ReplChars(FormData("SE"),"^","@")
   strEmailMsg = FormData("Msg")
   If FormData("BSC") = true or FormData("BSC") = "on" Then
      bolBldShortCut = true
   End If
   strBCSender = FormData("BCS")
   strTextMsg = FormData("TF")
   If FormData("RR") = true or FormData("RR") = "on" or FormData("RR") = "Y" Then
      bolRR = true
   Else
      bolRR = false
   End If
   strMsgFn = FormData("FN")

   If bolBcMailList and not bolIndivMailing Then
      strToEmail = strFrName & "<" & strFrEmail & ">"
   End If

   strTemp = ReplChars(FormData("AF"),",",";")
   strTemp = trim(ReplChars(strTemp,CHR(13)," "))
   If right(strTemp,1) = ";" Then
      strTemp = left(strTemp,len(strTemp)-1)
   End If

   If left(strTemp,1) = "?" Then
      If aryErrorText(1) = "" Then
         strTemp = REPLACE(strTemp,"/","\")
         aryErrorText(2) = " onload = " & CHR(34) & "CallDataPrompt('Attach File Lookup','Enter File Specification - Separate multiples with \'<font color=\'white\' size=+1><b>;</b></font>\'','\" & mid(strTemp,2) & "\')" & CHR(34)
      End If
      strTemp = ""
   ElseIf strTemp <> "" and left(strTemp,1) <> "/" and right(strTemp,1) <> "/" Then
      If left(strTemp,1) <> "\" Then
         strTemp = "\" & strTemp
      End If
      If right(strTemp,1) <> "\" Then
         strTemp = strTemp & "\"
      End If
   End If
   If left(strTemp,1) = "\" Then
      bolGetAttFile = true
   End If
   If FormData("CF") = "Y" Then
      'List Files Date Today Only
      strSelDate = left(NOW(),instr(NOW()," ")-1)
   End If
   intSub = 0
   do while strTemp <> ""
      intLoc = instr(strTemp,";")
      If intLoc=0 Then
         intLoc = len(strTemp)+1
      End If
      strHld = trim(left(strTemp,intLoc-1))
      If strHld <> "" Then
         intSub = intSub + 1
         If intSub > UBound(aryFileAtt) Then
            intSub = intSub -1
            EXIT DO
         End If
         If instr(strHld,"(") > 0 and instr(strHld,")") > 0 Then
            aryFileAtt(intSub) = trim(left(strHld,instr(strHld,"(")-1))
            strHld = mid(strHld,instr(strHld,"(")+1)
            strHld = trim(left(strHld,instr(strHld,")")-1))
         Else
            aryFileAtt(intSub) = strHld
         End If
         If instr(strHld,"*") = 0 Then
            aryFileAttAsIs(intSub) = GetFileName(strHld)
         End If
      End if
      strTemp = mid(strTemp,intLoc+1)
   Loop
   aryFileAtt(0) = intSub

   If aryFileAtt(0) > 0 or SESSION("AFH") <> "" Then
      bolAttFnOption = true
   End If

   If Request.Form("SEND") = "Preview" Then
      bolPreMsg = true
   End If
   If SESSION("EMPGM") <> "" Then
      Session.Contents.Remove("EMPGM")
   End If

   aryEM(1) = strToEmail
   aryEM(2) = strCCEmail
   aryEM(3) = strBCCEmail

   For intSub = 1 to 3
      aryEM(intSub) = ReplChars(aryEM(intSub),",",";")
      aryEM(intSub) = RemoveParmChars(aryEM(intSub))
      aryEM(intSub) = ReplChars(aryEM(intSub),"{","<")
      aryEM(intSub) = ReplChars(aryEM(intSub),"}",">")
      aryEM(intSub) = ReplChars(aryEM(intSub),"^","@")
   Next

   If SESSION("RQL") = "" Then
      SESSION("RQL") = 5
   End If

   strSpecSecGrp = "WebMail"
   Call Logon_Check(GetCurPath("")&strGoPgm,SESSION("RQL"),strLogonGrp)

   strGetData = SESSION("GETDATA")
   Session.Contents.Remove("GETDATA")

   If Request.QueryString("EMTP") <> "" Then
      SESSION("EMTP") = "." & Request.QueryString("EMTP")
   End If

   If SESSION("EMTP") = "" Then
      If strGetData = "Y" Then
         SESSION("EMTP") = Request.Form("EMTP")
      ElseIf Request.Form("Reset Form") <> "" Then
         strHld = SESSION("P")
         If strHld = "" Then
            strHld = strGoPgm
         End If
         SESSION("EMTP") = GetSelection(Application("EMTP"),"EMTP","Select Template Extension",strHld)
      End If
      If SESSION("EMTP") = "" Then
         SESSION("EMTP") = DefaultExt(Application("EMTP"))
      End If
      If left(SESSION("EMTP"),1) <> "." Then
         SESSION("EMTP") = "." & SESSION("EMTP")
      End If
      SESSION("EMTP") = ucase(SESSION("EMTP"))
   End If

   If right(ucase(strMsgFn),len(SESSION("EMTP"))) <> SESSION("EMTP") Then
      strMsgFn = strMsgFn & SESSION("EMTP")
   End If

   If SESSION("SECLOG") = true Then
      bolAppendFile = true
      aryTxtRecs(1) = NOW & ": " & SESSION("USERNAME") & " (" & SESSION("USEREMAIL") & ")"
      aryTxtRecs(0) = 1
      call Write_Text_File("Email.log",strFileLoc)
   End If

   Set errMsgs = CreateObject("Scripting.Dictionary")
   intErrNbr = 0

   If (Request.Form("SEND") = "" or strGetData = "Y") and not bolGetAttFile Then
      If strMsgFn <> "" Then
         call Process_Template
      End If
      If not bolGetAttFile Then
         If Request.Form("FormData") = "Y" and Request.Form("Reset Form") = "" Then
            ' Don't retrieve cookie values
         Else
            If trim(strFrName) = "" Then
               strFrName=Request.Cookies("email_"&SESSION("USERID"))("FrName")
            End If
            If trim(strFrEmail) = "" or strFrEmail = "?" Then
               If strFrEmail = "?" Then
                  strHld = "?"
               Else
                  strHld = ""
               End If
               strFrEmail=strHld & Request.Cookies("email_"&SESSION("USERID"))("FrAddr")
            End If
            If trim(strBCSender) = "" Then
               strBCSender=Request.Cookies("email_"&SESSION("USERID"))("BCS")
            End If
            If FormData("RR") = "" Then
               bolRR = Request.Cookies("email_"&SESSION("USERID"))("RR")
            End If
         End If
         Call Display_Form
         Response.End
      End If
   End If
   strHld = ""
   For intSub = 1 to aryFileAtt(0)
      If left(aryFileAtt(intSub),1) = "\" Then
         strHld = "AF"
         strFileAtt = aryFileAtt(intSub)
         If instr(strFileAtt,"*") = 0 and right(strFileAtt,1) <> "\" Then
            strFileAtt = strFileAtt & "\"
         End If
         aryFileAtt(intSub) = ""
         aryFileAttAsIs(intSub) = ""
         Call AddCookie("email_"&SESSION("USERID"), "AFSpec", strFileAtt, 365)
         EXIT FOR
      End If
   Next
   If strHld = "" and left(Request.Form("Send"),8) = "Template" Then
      strHld = "FN"
      strFileAtt = "\" & strFileLoc & "\*" & SESSION("EMTP")
   End If
   If strHld = "FN" Then
       strTitle = "Select Email Template"
   Else
       strTitle = "Select File(s) to Attach"
   End If
   If strHld <> "" Then
       intLoc = instr(strFileAtt,"*")
       If intLoc > 0 Then
          strFSpec = mid(strFileAtt,intLoc+1)
          strFileAtt = left(strFileAtt,intLoc-1)
       ElseIf left(strFileAtt,1) = "\" Then
          strFn = GetFileName(strFileAtt)
          If strFn <> "" Then
             strFSpec = strFn
             strFileAtt = strFilePath
          End If
       End If
       call Get_Folder_File_Details(strFileAtt,strFSpec)
       For intSub = 1 to aryFDtls(0,0)
         If trim(aryFDtls(1,intSub)) <> "" Then
            intCntr = intCntr + 1
            aryData(1,intCntr) = aryFDtls(1,intSub)
            If strHld = "FN" Then
                aryData(2,intCntr) = aryFDtls(1,intSub)
                aryData(3,intCntr) = "R"
            Else
                aryData(2,intCntr) = MapURL(aryFDtls(2,intSub))
                aryData(3,intCntr) = "C"
            End If
            aryData(4,intCntr) = strHld
         End If
       next
       For intSub = 1 to aryFileAtt(0)
          If trim(aryFileAtt(intSub)) <> "" Then
             intCntr = intCntr + 1
             aryData(1,intCntr) = aryFileAtt(intSub)
             If aryFileAttAsIs(intSub) <> "" Then
                aryData(1,intCntr) = aryData(1,intCntr) & "(" & aryFileAttAsIs(intSub) & ")"
             End If
             aryData(2,intCntr) = aryData(1,intCntr)
             aryData(3,intCntr) = "H"
             aryData(4,intCntr) = "AF"
          End If
       Next
       If Request.Form.Count >0 Then
          For Each X in Request.Form
             If X <> strHld Then
                SESSION(X) = Request.Form(X)
             End If
          Next
       ElseIf Request.QueryString.Count > 0 Then
          For Each X in Request.QueryString
             If X <> strHld Then
                SESSION(X) = Request.QueryString(X)
             End If
          Next
       End If
       SESSION("FN") = ""
       aryData(0,0) = intCntr
       strFldAlign = "R"
       call Get_Data(strGoPgm,strTitle,0)
       Response.End
   End If
   If right(ucase(trim(strSubj)),len(SESSION("EMTP"))) = SESSION("EMTP") and instr(trim(strSubj)," ") = 0 Then
      strMsgFn = trim(strSubj)
      strSubj = ""
   End If
   If bolBldShortCut and trim(strMsgFn) <> "" Then
      'Assume required elements in text file
   Else
      If Request.Form("SEND") = "Send" or bolPreMsg Then
         Call Validate_Input
      End If
   End if
   If errMsgs.Count > 0 Then
     call Process_Errors
   Else
     if intErrNbr = 0 then
        If bolBldShortCut Then
           intErrNbr = 0
        Else
           If strBCSender = "on" and not bolBldShortCut Then
              strSenderEM = strFrName & " <" & strFrEmail & ">"
              If aryEM(3) <> "" Then
                 strSenderEM = ";" & strSenderEM
              End If
           End If
           If bolPreMsg or Request.Form("SEND") = "Directory" or Request.Form("SEND") = "Attach Files" Then
              SESSION("BCS") = "off"
              SESSION("RR") = "off"
              SESSION("TF") = "off"
           End If
           For Each X in Request.Form
              If bolPreMsg or Request.Form("SEND") = "Directory" or Request.Form("SEND") = "Attach Files" Then
                 If X <> "Send" and X <> "BSC" Then
                    SESSION(X) = Request.Form(X)
                 End If
              Else
                 Session.Contents.Remove(X)
              End If
           Next
           If Request.Form("SEND") = "Directory" Then
              SESSION("EMPGM") = "email.asp"
              SESSION("emKW") = ""
              If left(strSubj,1) = "?" Then
                 SESSION("emKW") = mid(strSubj,2)
                 SESSION("SUBJ") = ""
              End If
              If SESSION("emKW") = "" Then
                 For intSub = 1 to 3
                    If left(aryEM(intSub),1) = "?" Then
                       SESSION("emKW") = mid(aryEM(intSub),2)
                       If intSub = 1 Then
                          SESSION("TO") = ""
                       ElseIf intSub = 2 Then
                          SESSION("CC") = ""
                       ElseIf intSub = 3 Then
                          SESSION("BCC") = ""
                       End If
                       EXIT FOR
                    End If
                 Next
              End If
              If SESSION("emKW") = "" and left(strEmailMsg,1) = "?" Then
                 SESSION("emKW") = mid(strEmailMsg,2)
                 SESSION("MSG") = ""
              End If
              SESSION("TO") = ReplChars(SESSION("TO"),"<","{")
              SESSION("TO") = ReplChars(SESSION("TO"),">","}")
              SESSION("CC") = ReplChars(SESSION("CC"),"<","{")
              SESSION("CC") = ReplChars(SESSION("CC"),">","}")
              SESSION("BCC") = ReplChars(SESSION("BCC"),"<","{")
              SESSION("BCC") = ReplChars(SESSION("BCC"),">","}")
              If SESSION("KW") = "" Then
                 SESSION("KW") = BuildKW(SESSION(""))
              End If
              Response.Redirect("build_email_list.asp")
              Response.End
           End If
           If Request.Form("SEND") = "Attach Files" Then
              SESSION("EMPGM") = "email.asp?AF="
              SESSION("AFH") = Request.Form("AF")
              Response.Redirect("upload.asp")
              Response.End
           End If
           If not bolBldShortCut Then
              Call AddCookie("email_"&SESSION("USERID"), "FrName", strFrName, 365)
              Call AddCookie("email_"&SESSION("USERID"), "FrAddr", strFrEmail, 365)
              Call AddCookie("email_"&SESSION("USERID"), "BCS", strBCSender, 365)
              Call AddCookie("email_"&SESSION("USERID"), "TF", strTextMsg, 365)
              Call AddCookie("email_"&SESSION("USERID"), "RR", bolRR, 365)
           End If
           SESSION("TO") = aryEM(1)
           strFileAtt = ""
           strFileAttAsIs = ""
           If bolIndivMailing and (aryEM(2) <> "" or aryEM(3) <> "") Then
              bolIndivMailing = false
           End If
           For intSub = 1 to aryFileAtt(0)
              If left(aryFileAtt(intSub),1) = "/" Then
                 aryFileAtt(intSub) = "http://" & Request.ServerVariables("SERVER_NAME") & aryFileAtt(intSub)
              End If
              If strFileAtt = "" Then
                 strFileAtt = aryFileAtt(intSub)
                 strFileAttAsIs = aryFileAttAsIs(intSub)
              Else
                 strFileAtt = strFileAtt & ";" & aryFileAtt(intSub)
                 strFileAttAsIs = strFileAttAsIs & ";" & aryFileAttAsIs(intSub)
              End If
           Next
           strFileAtt = ReplParmChars(strFileAtt)
           If strTextMsg = "on" Then
              strEmailMsg = "<TEXT>" & strEmailMsg
           End If
           If bolIndivMailing Then
              intSub = 1
              bolDoPhrases = false
              strWrdSep = ";"
              strEM = trim(GetWords(aryEM(1),intSub,1))
           Else
              strEM = trim(aryEM(1))
           End If
           If right(strEM,1) = ";" Then
              strEM = left(strEM,len(strEM)-1)
           End If
           If bolIndivMailing Then
              bolGetMbrDetails = true
           End If
           do while strEM <> ""
              If strDupeChkEmail = "" and instr(strEM,";") = 0 and APPLICATION("DIRTBL") <> "" Then
                 strDupeChkEmail = strEM
              End If
              Call Load_Variables(strEM)
              strMsg = CheckTextVariables(strEmailMsg)
              intErrNbr = SendJMail(strSMTPServer,strEM,aryEM(2),aryEM(3)&strSenderEM,strFrName,strFrEmail,strSubj,strMsg,strFileAtt,strFileAttAsIs,bolRR)
              If bolIndivMailing Then
                 bolDoPhrases = false
                 strWrdSep = ";"
                 intSub = intSub + intRetWords
                 strEM = GetWords(aryEM(1),intSub,1)
              Else
                 strEM = ""
              End If
           loop
        End If

        If intErrNbr = 0 Then
           bolNoTable = true
           Call Setup_Web_Page("","")
           Response.Write "<table>" & vbCrLf
           If not bolPreMsg Then
              If bolBldShortCut Then
                 strPgm = "http://" & GetPgmPath("") & strGoPgm
                 If strFrName <> "" Then
                    strPgm = strPgm & "?SN=" & strFrName & "&SE=" & ReplChars(strFrEmail,"@","^")
                 End If
                 strPgm = ReplParmChars(strPgm)
                 If strTextMsg = "on" then
                    strPgm = strPgm & "&TF=Y"
                 End If
                 If bolRR Then
                    strPgm = strPgm & "&RR=Y"
                 End If
                 If strBCSender = "on" Then
                    strPgm = strPgm & "&BCS=Y"
                 End If
              Else
                 Response.Write "<tr><td></td><td><b>Email Sucessfully Sent</b><br><br></td></tr>" & vbCrLf
                 strPgm = ""
              End If
           End If
           If bolBcMailList Then
              aryEM(1) = "Mailing List"
              aryEM(2) = ""
              aryEM(3) = ""
           End If
           For intSub = 1 to 3
              If bolBldShortCut Then
                 If intSub = 1 Then
                    strHdr = "&TO="
                 ElseIf intSub = 2 Then
                    strHdr = "&CC="
                 Else
                    strHdr = "&BCC="
                 End If
              Else
                 If intSub = 1 Then
                    strHdr = "To:"
                 ElseIf intSub = 2 Then
                    strHdr = "CC:"
                 Else
                    strHdr = "BCC:"
                 End If
                 strHdr = "<font color='red'>" & strHdr & " &nbsp;&nbsp;</font>"
              End If
              strTemp = aryEM(intSub)
              do while strTemp <> ""
                 intLoc = instr(strTemp,",")
                 If intLoc = 0 Then
                    intLoc = instr(strTemp,";")
                 End If
                 If intLoc = 0 or bolBldShortCut Then
                    intLoc = len(strTemp) + 1
                 End If
                 strEM = trim(left(strTemp,intLoc-1))
                 If strEM <> "" and bolBldShortCut Then
                    strEM = ReplParmChars(strEM)
                    strEM = ReplChars(strEM,"<","{")
                    strEM = ReplChars(strEM,">","}")
                    strEM = ReplChars(strEM,"@","^")
                    strPgm = strPgm & strHdr & strEM
                    strHdr = ""
                 End If
                 If strEM <> "" and not bolBldShortCut and not bolPreMsg Then
                    Response.Write "<tr><td nowrap align='right'>" & strHdr & "</td><td><b>" & ReplBrackets(strEM) & "</b></td></tr>"  & vbCrLf
                 End If
                 strHdr = ""
                 strTemp = mid(strTemp,intLoc+1)
              loop
           Next
           If bolBldShortCut Then
              Response.Write "<tr><td></td><td><b><font color='red'>Copy and Paste Link to Shortcut Entry</font></b><br><br></td></tr>" & vbCrLf
              Response.Write "<tr><td></td><td><b>" & ReplBrackets(strPgm) & "</b></td></tr>"  & vbCrLf
           End If
           Response.Write "<tr><td align='center' colspan='2'><br><br>" & vbCrLf
           Response.Write "<input type=" & CHR(34) & "button" & CHR(34) & " value=" & CHR(34) & "Back" & CHR(34) & " onClick=" & CHR(34) & "document.location.href = '" & strGoPgm & "'" & CHR(34) & ">" & vbCrLf
           If instr(strDupeChkEmail,";") = 0 and APPLICATION("DIRTBL") <> "" Then
              If not EmailDuplicate(strDupeChkEmail,"") Then
                 strHld="BPG=1&AC=A&IA=Y&email="&strDupeChkEmail&"&Last%20Name="&strLName&"&First%20Name="&strFName
                 Response.Write "&nbsp;&nbsp;&nbsp;" & vbCrLf
                 Response.Write "<input type=" & CHR(34) & "button" & CHR(34) & " value=" & CHR(34) & "Update Directory" & CHR(34) & " onClick=" & CHR(34) & "document.location.href = 'directory_maint.asp?" & strHld & "'" & CHR(34) & ">" & vbCrLf
              End If
           End If
           Response.Write "</td></tr>" & vbCrLf
           call Wrapup_Web_Page
        End If
     End If
     if intErrNbr <> 0 then
        aryErrorMsgs = errMsgs.Items
        call Process_Errors
     End If
   End If

   If bolEditError Then
      Call Display_Form
   End If

SUB Validate_Input

  '// Validate form entries.

   DIM strHld

   intErrNbr = 0

   If bolPreMsg Then
      If strTextMsg = "on" Then
         intErrNbr = intErrNbr + 1
         errMsgs.Add intErrNbr, "Preview Not Available for Text Format Messages"
      ElseIf trim(strEmailMsg) = "" Then
         intErrNbr = intErrNbr + 1
         errMsgs.Add intErrNbr, "Message Text Required for Preview Function"
      End If
   End If
   If Request.Form("SEND") = "Send" Then
      If trim(strFrName) = "" Then
         intErrNbr = intErrNbr + 1
         errMsgs.Add intErrNbr, "Sender Name is Required"
      End If
      strHld = strFrEmail
      If left(strHld,1) = "-" Then
         strHld = mid(strHld,2)
      End If
      If trim(strHld) = "" Then
         intErrNbr = intErrNbr + 1
         errMsgs.Add intErrNbr, "Sender Email Address is Required"
      ElseIf ValidateEMailFormat(LCase(strHld)) = false Then
         intErrNbr = intErrNbr + 1
         errMsgs.Add intErrNbr, "Invalid Sender E-Mail Address Format"
      End If
      If trim(aryEM(1)) = "" Then
         intErrNbr = intErrNbr + 1
         errMsgs.Add intErrNbr, "At least one TO Email Address is Required"
      End If
      If trim(strSubj) = "" Then
         intErrNbr = intErrNbr + 1
         errMsgs.Add intErrNbr, "Message Subject is Required"
      End If
      For intSub = 1 to aryFileAtt(0)
         If left(ucase(aryFileAtt(intSub)),5) <> "HTTP:" Then
            If CheckFileExist(aryFileAtt(intSub)) > 0 Then
               intErrNbr = intErrNbr + 1
               errMsgs.Add intErrNbr, "File Attachment " & aryFileAtt(intSub) & " Not Found On Server"
            End If
         End If
      Next
      If trim(strEmailMsg) = "" Then
         strEmailMsg = " "
      End If
   End If

   If errMsgs.Count > 0 Then
      aryErrorMsgs = errMsgs.Items
   End If
   err.clear

End SUB

FUNCTION DefaultExt(EXT)

    DIM intSub, strWord, strExt

    DefaultExt = ".EM"

    strExt = REPLACE(EXT,";"," ")

    bolDoPhrases = false
    intSub = 1
    strWord = GetWords(strExt,intSub,1)
    do while strWord <> ""
       strWord = trim(REPLACE(strWord,"."," "))
       If FileCount("\" & strFileLoc & "\*." & strWord) > 0 Then
          DefaultExt = "." & ucase(strWord)
          EXIT DO
       End If
       intSub = intSub + 1
       strWord = GetWords(strExt,intSub,1)
    loop

END FUNCTION

SUB Process_Template

   DIM intAttNbr

   intAttNbr = aryFileAtt(0)

   call Read_Text_File(strMsgFn,strFileLoc)
   For intSub = 1 to aryTxtRecs(0)
      If left(aryTxtRecs(intSub),1) <> ":" Then
         strHld = trim(mid(aryTxtRecs(intSub),4))
         If left(ucase(aryTxtRecs(intSub)),3) = "SB:" Then
            If trim(FormData("Subj")) = "" Then
               strSubj = strHld
               SESSION("Subj") = strHld
            End If
         ElseIf left(ucase(aryTxtRecs(intSub)),3) = "SN:" Then
            strFrName = strHld
         ElseIf left(ucase(aryTxtRecs(intSub)),3) = "SE:" Then
            strFrEmail = strHld
         ElseIf left(ucase(aryTxtRecs(intSub)),3) = "TO:" Then
            If trim(FormData("TO")) = "" Then
               aryEM(1) = strHld
               SESSION("TO") = strHld
            End If
         ElseIf left(ucase(aryTxtRecs(intSub)),3) = "CC:" Then
            If trim(FormData("CC")) = "" Then
               aryEM(2) = strHld
               SESSION("CC") = strHld
            End If
         ElseIf left(ucase(aryTxtRecs(intSub)),4) = "BCC:" Then
            If trim(FormData("BCC")) = "" Then
               aryEM(3) = trim(mid(aryTxtRecs(intSub),5))
               SESSION("BCC") = strHld
            End If
         ElseIf left(ucase(aryTxtRecs(intSub)),3) = "AF:" Then
             intAttNbr = intAttNbr + 1
             aryFileAtt(intAttNbr) = strHld
             aryFileAttAsIs(intAttNbr) = GetFileName(strHld)
             bolAttFnOption = true
             aryFileAtt(0) = intAttNbr
             If left(aryFileAtt(intAttNbr),1) = "\" Then
                bolGetAttFile = true
             End If
         ElseIf left(ucase(aryTxtRecs(intSub)),4) = "BCS:" Then
            strBCSender = mid(ucase(aryTxtRecs(intSub)),5,1)
            SESSION("BCS") = strBCSender
         ElseIf left(ucase(aryTxtRecs(intSub)),3) = "TF:" Then
            strTextMsg = mid(ucase(aryTxtRecs(intSub)),4,1)
            SESSION("TF") = strTextMsg
         ElseIf left(ucase(aryTxtRecs(intSub)),3) = "RR:" Then
            If mid(ucase(aryTxtRecs(intSub)),4,1) = "Y" Then
               bolRR = true
            Else
               bolRR = false
            End If
            SESSION("RR") = bolRR
         Else
            strEmailMsg = strEmailMsg & aryTxtRecs(intSub) & CHR(10)
            SESSION("MSG") = strEmailMsg
         End If
      End If
   Next

END SUB

SUB Display_Form

   DIM strT


   If strBCSender = "Y" Then
      strBCSender = "on"
   ElseIf strBCSender = "N" Then
      strBCSender = "off"
   End If

   If strTextMsg = "Y" Then
      strTextMsg = "on"
   ElseIf strTextMsg = "N" Then
      strTextMsg = "off"
   End If

   strFileAtt = ""
   For intSub = 1 to aryFileAtt(0)
      aryFileAtt(intSub) = aryFileAtt(intSub)
      If aryFileAttAsIs(intSub) <> "" Then
         aryFileAtt(intSub) = aryFileAtt(intSub) & "(" & aryFileAttAsIs(intSub) & ")"
      End If
      If strFileAtt = "" Then
         strFileAtt = aryFileAtt(intSub)
      Else
         strFileAtt = strFileAtt & "; " & aryFileAtt(intSub)
      End If
   Next
   If SESSION("AFH") <> "" THen
      strFileAtt = SESSION("AFH") & "; " & strFileAtt
      Session.Contents.Remove("AFH")
   End If
%>

   <html>

   <head>
   <title>Send Web Mail</title>
   <%
   call SetUp_Data_Prompt
   strAFSpec = InsertDoubleBS(Request.Cookies("email_"&SESSION("USERID"))("AFSpec"))
   %>
   <SCRIPT language='javascript'>

   function GetAttachFileSpec(value) {

      if(value.length<=0)
      	return false;
      else
        X=document.EmailForm
      	X.AF.value = value+";"+X.AF.value
           document.EmailForm.submit()
   }
   </SCRIPT>

   <%Check_For_PopUp_Msg%>
   </head>

   <%Response.Write aryErrorText(1) & vbCrLf%>
   <body bgcolor="#008080"<%=aryErrorText(2)%>>
   <%Response.Write aryErrorText(3) & vbCrLf%>

   <div align="CENTER">
   <table width="500" border="1" cellspacing="0" cellpadding="2" bordercolor="#000080">

      <tr>
        <td width="50%" align="center" bgcolor="#000095">
        <Font color="white">
        <b>Send Web Mail</b>
        </font></td>
      </tr>

      <tr>
        <td valign="TOP" bgcolor="#B4B4B4">
        <font face="Arial" size="2">

     <div align="CENTER">
     <table width="90" border="0" cellspacing="0" cellpadding="1">
        <form Name="EmailForm" Action="<%=strGoPgm%>" Method="post">
        <input type="Hidden" name="FormData" value="Y">
       <tr>
 	  <td nowrap align="right">
          <%If trim(strFrName) = "" Then
               strHld = " color='red'"
            Else
               strHld = ""
            End If%>
         <font size="-1"<%=strHld%>> <b>Name</b></font>&nbsp;&nbsp;</td>
          <td nowrap>
          <input type="Text" name="SN" size="55" value="<%=strFrName%>">
          <%If trim(strFrEmail) = "" Then
               strHld = " color='red'"
            Else
               strHld = ""
            End If%>
          &nbsp;&nbsp;<font size="-1"<%=strHld%>><b>e-Mail</b></font>&nbsp;
          <input type="Text" name="SE" size="72" value="<%=strFrEmail%>">
          </td>
        </tr>
        <tr>
        <td nowrap align="right">
          <%If trim(aryEM(1)) = "" Then
               strHld = " color='red'"
            Else
               strHld = ""
            End If%>
       <font size="-1"<%=strHld%>><b>To</b></font>&nbsp;&nbsp;
        </td>
        <td nowrap>
        <%
        intRows = Round((len(aryEM(1))/107)+.5)
        If intRows > 10 Then
           intRows = 10
        End if
        If bolBcMailList Then%>
           <input type="Text" readonly size="141" value="Mailing List">
           <%If bolIndivMailing Then%>
           <input type="hidden" name="TO" value="<%=ReplChars(SESSION("ML"),",",";") %>">
           <%Else%>
           <input type="hidden" name="BCC" value="<%=ReplChars(SESSION("ML"),",",";") %>">
           <%End IF%>
        <%ElseIf intRows < 2 Then%>
           <input type="Text" name="TO" size="141" value="<%=ReplChars(aryEM(1),",",";") %>">
        <%Else%>
           <textarea name="TO" cols="106" rows="<%=intRows%>" wrap="YES"><%=ReplChars(aryEM(1),",",";") %></textarea>
        <%End If%>
        </td>
        </tr>
        <%If not bolBcMailList Then%>
        <tr>
        <td nowrap align="right">
       <font size="-1"> <b>CC</b></font>&nbsp;&nbsp;
        </td>
        <td nowrap>
        <%
        intRows = Round((len(aryEM(2))/107)+.5)
        If intRows > 10 Then
           intRows = 10
        End if
        If intRows < 2 Then%>
           <input type="Text" name="CC" size="141" value="<%=ReplChars(aryEM(2),",",";") %>">
        <%Else%>
           <textarea name="CC" cols="106" rows="<%=intRows%>" wrap="YES"><%=ReplChars(aryEM(2),",",";") %></textarea>
        <%End If%>
        </td>
        </tr>
        <%IF trim(aryEM(3)) <> "" Then%>
        <tr>
        <td nowrap align="right">
       <font size="-1"> <b>BCC</b></font>&nbsp;&nbsp;
        </td>
        <td nowrap>
        <%
        intRows = Round((len(aryEM(3))/107)+.5)
        If intRows > 10 Then
           intRows = 10
        End if
        If intRows < 2 Then%>
           <input type="Text" name="BCC" size="141" value="<%=ReplChars(aryEM(3),",",";") %>">
        <%Else%>
           <textarea name="BCC" cols="106" rows="<%=intRows%>" wrap="YES"><%=ReplChars(aryEM(3),",",";") %></textarea>
        <%End If%>
        </td></tr>
        <%End If%>
        <%End If%>
        <tr>
        <td nowrap align="right">
          <%If trim(strSubj) = "" Then
               strHld = " color='red'"
            Else
               strHld = ""
            End If%>
        <font size="-1"<%=strHld%>><b>Subject</b></font>&nbsp;&nbsp;
        </td>
        <td nowrap>
        <input type="Text" name="Subj" size="141" value="<%=strSubj %>">
        </td>
        </tr>
        <%IF bolAttFnOption Then%>
        <tr>
        <td nowrap align="right">
        <font size="-1"> <b>Attach..</b></font>&nbsp;&nbsp;
        </td>
        <td nowrap>
        <%
        intRows = Round((len(strFileAtt)/107)+.5)
        If intRows > 10 Then
           intRows = 10
        End if
        If intRows < 2 Then%>
           <input type="text" name="AF" size="141" value="<%=strFileAtt %>">
        <%Else%>
           <textarea name="AF" cols="106" rows="<%=intRows%>" wrap="YES"><%=strFileAtt %></textarea>
        <%End If%>
        </td></tr>
        <%Else%>
           <input type="hidden" name="AF" size="141" value="">
        <%End If%>
        <tr>
        <td align="right" colspan="2">
        <%If strBCSender = "on" Then
           strBCSender = "checked"
        Else
           strBCSender = ""
        End If%>
        <input type="Checkbox" name="BCS" <%=strBCSender%>>
        <font size="-1">
        <b>Bcc Sender</b>
        </FONT>
        <%If bolRR Then
           strHld = "checked"
        Else
           strHld = ""
        End If%>
        <input type="Checkbox" name="RR" <%=strHld%>>
        <font size="-1">
        <b>Return Receipt</b>
        </FONT>
        <%If strTextMsg = "on" Then
           strTextMsg = "checked"
        Else
           strTextMsg = ""
        End If%>
        <%If not bolBcMailList Then%>
        <input type="Checkbox" name="TF" <%=strTextMsg%>>
        <font size="-1">
        <b>Text Format</b>
        </FONT>
        <input type="Checkbox" name="BSC">
        <font size="-1">
        <b>Shortcut</b>
        </FONT>
        <%End If%>
        &nbsp;&nbsp;
        <br>
        <%If strBrowser = "Netscape" Then
             intRows = 18
          Else
             intRows = 20
          End If
          If trim(SESSION("BCC"))="" Then
             intRows = intRows + 2
          End If
          IF bolAttFnOption Then
             intRows = intRows - 2
          End If
          %>
        <div align='left'>
        <textarea name="Msg" cols="114" rows="<%=intRows%>" wrap="YES"><%=strEmailMsg %></textarea>
        </div>
        </td>
        </tr>
        <tr>
        <td colspan="2" nowrap align="center">
        <hr width="900" size="2" color="#000000" noshade>
        </td>
        </tr>
        <tr>
        <td nowrap align="center" colspan="2">
        <%
        IF SESSION("NBRTP") = "" Then
           SESSION("NBRTP") = FileCount("\" & strFileLoc & "\*" & SESSION("EMTP"))
        End If
        strHld = "&nbsp;&nbsp;&nbsp;"
        %>
        <%If bolGetAttFile Then%>
        <input type="Submit" name="Send" value="Select Attachment">
        <%Else%>
        <input type="Submit" name="Send" value="Send">
        <%End If%>
        <%If not bolGetAttFile Then%>
        <%If not bolBcMailList and APPLICATION("DIRTBL") <> "" Then%>
        <%=strHld%>
        <input type="Submit" name="Send" value="Directory">
        <%End If%>
        <%=strHld%>
        <input type="Submit" name="Send" value="Preview">
        <%End If%>
        <%=strHld%>
        <%strT = "Template"
          If instr(Application("EMTP"),";") > 0 Then
             strT = strT & " (" & lcase(mid(SESSION("EMTP"),2)) & ")"
          End If
          If SESSION("NBRTP") > 0 Then %>
        <input type="Submit" name="Send" value="<%=strT%>">
        <%Else%>
        <input type="Text" value="<%=strT%>" size=8>
        <%End If%>
        <%=strHld%>
        <input type="Submit" name="Send" value="Attach Files">
        <%=strHld%>
        <input type="Submit" Name="Reset Form" value="Clear">
        <%If not bolBcMailList Then%>
        <%=strHld%>
        <input type="button" value="Help" onClick="window.open('Send Web Mail Documentation.htm')">
        <%End If%>
        </td>
        </tr>
        </form>
        </table>
        </table>
	</div>
   </div>
   </body>
   </html>

<% End SUB %>


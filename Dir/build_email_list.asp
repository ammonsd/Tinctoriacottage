<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>
<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/popup_email.inc" -->
<!--#INCLUDE FILE="../dir/maint_webpage.inc" -->
<!--#INCLUDE FILE="../dir/display_edit_errors.inc" -->
<!--#INCLUDE FILE="../dir/cookie_maint.inc" -->
<!--#INCLUDE FILE="../dir/security_select.inc" -->
<!--#INCLUDE FILE="../dir/directory_select.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->
<!--#INCLUDE FILE="../dir/directory_navigation_details.inc" -->
<%

   DIM intSub, strWord, strHld, intTotSel, strLastEmUsed, strLastNameUsed, strPH, intCols, strSecEM, bolNPOption, strGoTo
   DIM intMaxCols

   strGoTo = "build_email_list.asp"

   strSpecOpt = Request.QueryString("SO")
   If strSpecOpt = "" Then
      strSpecOpt = Request.Form("SO")
   End If
   strSecEM = Request.Form("SECUR")
   If strSecEM = "Y" Then
      If Request.Form("PROCESS") <> "Y" Then
         strTempKW = Request.Form("KW")
         SESSION("SRCHKW") = ""
      End If
      bolSesKWOnly = true
   End If
   If SESSION("EMPGM") <> "" and SESSION("emKW") <> "" Then
      strSpecOpt = "Y"
   End If

   If Request.QueryString("NP") = "Y" and trim(SESSION("TO")) <> "" Then
      bolNPOption = true
   End If

   If bolNPOption Then
      SESSION("VCF") = "H"
   Else
      SESSION("VCF") = Request.QueryString("V")
   End If

   If Request.QueryString("V") = "H" and (Request.Form("GETKW") = "Y" or SESSION("DIRADDR") = "Y") Then
      Session.Contents.Remove("VCF-KYWRDS")
      Session.Contents.Remove("DIRADDR")
   End If

   If SESSION("VCF") = "" Then
      SESSION("VCF") = Request.Form("V")
   End If

   If SESSION("VCF") = "Y" or SESSION("VCF") = "H" Then
      Session.Contents.Remove("EMPGM")
      Session.Contents.Remove("EMN")
      If not bolNPOption Then
         Session.Contents.Remove("TO")
      End If
      Session.Contents.Remove("CC")
      Session.Contents.Remove("BCC")
   Else
      Session.Contents.Remove("VCF")
   End If

   If SESSION("RQL") = "" Then
      If SESSION("VCF") = "Y" Then
         SESSION("RQL") = 5
      Else
         SESSION("RQL") = 3
      End If
   End If

   If SESSION("VCF") = "Y" Then
      strSpecSecGrp = "VCF"
   End If

   Call System_Setup(strGoTo)
   Call Logon_Check(GetCurPath("")&strGoTo,SESSION("RQL"),strLogonGrp)

   Call Database_Setup

   If SESSION("KW") = "" and strTempKW = "" Then
      SESSION("KW") = BuildKW(SESSION("SECKW"))
   End If

   strTypeGoto = "NF"

   intMaxCols = 5

   If Request.Form("ALL") = "" and Request.Form("NONE") = "" and not bolNPOption Then
      If Request.Form("EMN") <> "" Then
         SESSION("EMN") = SESSION("EMN") & Request.Form("EMN") & " "
      End If
      If Request.Form("EM") <> "" Then
         If Request.Form("CC") <> "" Then
            strHld = "CC"
            strSpecOpt = "Y"
         ElseIf Request.Form("BCC") <> "" Then
            strHld = "BCC"
            strSpecOpt = "Y"
         Else
            strHld = "TO"
         End If
         If SESSION(strHld) = "" Then
            SESSION(strHld) = Request.Form("EM")
         Else
            SESSION(strHld) = SESSION(strHld) & "," & Request.Form("EM")
         End If
      End If
      If Request.Form("PROCESS") <> "Y" Then
          Call Verify_Address_Variables("TO")
          Call Verify_Address_Variables("CC")
          Call Verify_Address_Variables("BCC")
      End If
   End If

   If Request.Form("BUILD") <> "" or Request.Form("RETURN") <> "" or bolNPOption Then
      'Build Mailing List
      If SESSION("EMPGM") = "" and Request.Form("EM") = "" and SESSION("TO") = "" and SESSION("CC") = "" and SESSION("BCC") = "" and not bolNPOption Then
         Set errMsgs = CreateObject("Scripting.Dictionary")
         errMsgs.Add 1, "At Least One Email Address Must Be Selected"
         aryErrorMsgs = errMsgs.Items
         call Display_Errors
      End If
   Else
      If Request.Form("SEARCH") <> "Y" and Request.Form("ALL") = "" and Request.Form("NONE") = "" and strSpecOpt <> "Y"  Then
         If SESSION("EMPGM") = "" Then
            Session.Contents.Remove("EMN")
            Session.Contents.Remove("TO")
            Session.Contents.Remove("CC")
            Session.Contents.Remove("BCC")
         End If
         If SESSION("VCF") <> "" Then
            aryHidOpts(1,1) = "V"
            aryHidOpts(1,2) = SESSION("VCF")
            aryHidOpts(0,0) = 1
         End If
         If Request.Form("SEARCH") <> "Y" Then
            SESSION("AS") = "Y" 'Turn on Advance Search Options
            If Request.QueryString("SRCH") = "Y" Then
               bolSrchAryOnly = true
               Call Check_Search_Criteria("Y")
            End If
            SESSION("BC") = ""
            call Setup_Web_Page("Enter Search Criteria",2)
            Response.Write "<tr><td>" & vbCrLf
            intSrchRows = 5
            bolBottomButton = true
            Call Insert_Search_Prompt("build_email_list.asp")
            Response.Write "</td></tr>" & vbCrLf
            If APPLICATION("DIRTBL") <> "" Then
               Call Insert_Navigation_Details("S")
            End If
            call Wrapup_Web_Page
            Response.End
         End If
      Else
         If Request.Form("SEARCH") = "Y" Then
            Session.Contents.Remove("EMN")
            Session.Contents.Remove("TO")
            Session.Contents.Remove("CC")
            Session.Contents.Remove("BCC")
         End If
         If strSecEM = "Y" Then
            strTempKW = BuildSrchString(strTempKW)
         End If
         If SESSION("VCF") = "H" and SESSION("VCF-KYWRDS") <> "" Then
            SESSION("emKW") = SESSION("VCF-KYWRDS")
         End If
         If SESSION("VCF") = "H" and SESSION("VCF-KYWRDS") = "" and SESSION("SRCHKW") <> "" Then
            SESSION("VCF-KYWRDS") = SESSION("SRCHKW")
         End If
         If SESSION("emKW") <> "" Then
            SESSION("SRCHKW") = SESSION("emKW")
            Session.Contents.Remove("emKW")
         End If
         If strSecEM = "Y" Then
            SESSION("PTYP") = "SE"
            Call Build_Where_SQL
         Else
            Call Check_Search_Criteria("Y")
         End If
         call Setup_Web_Page("Select Members to be Included",intMaxCols)
         Call Get_EM_Addr
         Response.End
      End If
      If intTotSel = 0 and SESSION("EMPGM") <> "" and Request.Form("PROCESS") = "Y" Then
         Response.Clear
      Else
         call Wrapup_Web_Page
         Response.End
      End If
   End If

   If SESSION("TO") = "" Then
      SESSION("TO") = Request.Form("EM")
   End If

   If SESSION("VCF") <> "" Then
      strToAddr = trim(ReplChars(SESSION("TO"),","," "))
      Call Build_VCF_HTML_Entries
      Response.End
   End If

   strToAddr = trim(ReplChars(SESSION("TO"),",",";"))
   strCCAddr = trim(ReplChars(SESSION("CC"),",",";"))
   strBCcAddr = trim(ReplChars(SESSION("BCC"),",",";"))

   If SESSION("EMPGM") = "" Then
      call Setup_Web_Page("",2)
      If strToAddr <> "" Then
         Response.Write "<tr><td valign='top'>" & vbCrLf
         Response.Write "<b>To:</b>&nbsp;&nbsp;" & vbCrLf
         Response.Write "</td>" & vbCrLf
         Response.Write "<td width=700>" & vbCrLf
         Response.Write "<font color='red'>" & strToAddr & "</font><br>" & vbCrLf
         Response.Write "</td></tr>" & vbCrLf
      End If
      If strCCAddr <> "" Then
         Response.Write "<tr><td>&nbsp;</td></tr>" & vbCrLf
         Response.Write "<tr><td valign='top'>" & vbCrLf
         Response.Write "<b>CC:</b>&nbsp;&nbsp;" & vbCrLf
         Response.Write "</td>" & vbCrLf
         Response.Write "<td width=700>" & vbCrLf
         Response.Write "<font color='red'>" & strCCAddr & "</font><br>" & vbCrLf
         Response.Write "</td></tr>" & vbCrLf
      End If
      If strBCcAddr <> "" Then
         Response.Write "<tr><td>&nbsp;</td></tr>" & vbCrLf
         Response.Write "<tr><td valign='top'>" & vbCrLf
         Response.Write "<b>BCC:</b>&nbsp;&nbsp;" & vbCrLf
         Response.Write "</td>" & vbCrLf
         Response.Write "<td width=700>" & vbCrLf
         Response.Write "<font color='red'>" & strBCcAddr & "</font><br>" & vbCrLf
         Response.Write "</td></tr>" & vbCrLf
      End If
      Response.Write "<tr><td><br></td></tr>" & vbCrLf
      Response.Write "<tr><td nowrap align='center' colspan=2><hr width=700 size='5' noshade></td></tr>" & vbCrLf
      Response.Write "<tr><td colspan=2>" & vbCrLf
      If len(strToAddr) > 2000 Then
         strToAddr = "**Copy/Paste Address String Here**"
      End If
      If len(strCCAddr) > 2000 Then
         strCCAddr = "**Copy/Paste Address String Here**"
      End If
      If len(strBCcAddr) > 2000 Then
         strBCcAddr = "**Copy/Paste Address String Here**"
      End If
   Else
      bolNoDisplay = true
   End If
   Call Popup_eMail(strToAddr,strCCAddr,strBCcAddr,"","")
   If SESSION("EMPGM") <> "" Then
      Response.Redirect(SESSION("EMPGM"))
      Response.END
   End If
   Response.Write "</td></tr>" & vbCrLf
   Response.Write "<tr><td nowrap align='center' colspan=2><b>OR</b><br><br></td></tr>" & vbCrLf
   If strCCAddr <> "" or strBCcAddr <> "" Then
      strHld = "Copy/Paste Address Strings into Appropriate Email Address Areas"
   Else
      strHld = "Copy/Paste Address String into Email Address Area"
   End If
   Response.Write "<tr><td nowrap align='center' colspan=2><b>" & strHld & "</b><br><br></td></tr>" & vbCrLf
   strMaint = "Y"
   call Wrapup_Web_Page

SUB Get_EM_Addr

    DIM  strChk, intCol, strHld, strTemp

'    If Request.Form("ALL") <> "" or strSecEM = "Y" or Request.Form("PROCESS") <> "Y" Then
    If Request.Form("ALL") <> "" or SESSION("VCF") = "H" Then
       strChk = "checked"
    End If

    If strSecEM = "Y" Then
       Call Get_Security_Members
    Else
       bolBuildAryOnly = true
       Call Directory_Selection_List("")
    End If

    CALL Setup_ShowDetails_Script

    Response.Write "<form action='build_email_list.asp' method='post'>" & vbCrLf
    Response.Write "<input type='hidden' name='PROCESS' value='Y'>" & vbCrLf

    If Request.Form("Keywords") <> "" Then
       Response.Write "<input type='hidden' name='Keywords' value='" & Request.Form("Keywords") &"'>" & vbCrLf
    End If
    If SESSION("VCF") <> "" Then
       Response.Write "<input type='hidden' name='V' value='" & SESSION("VCF") &"'>" & vbCrLf
    End If
    If Request.QueryString("ST") = "Y" Then
       Response.Write "<input type='hidden' name='KWOVR' value='Y'>" & vbCrLf
    End If
    If strSpecOpt <> "" Then
       Response.Write "<input type='hidden' name='SO' value='" & strSpecOpt & "'>" & vbCrLf
    End If
    If strSecEM <> "" Then
       Response.Write "<input type='hidden' name='SECUR' value='" & strSecEM & "'>" & vbCrLf
    End If
    Response.Write "<tr>" & vbCrLf
    intCol = 0
    intTotSel = 0

    If strChk = "" and SESSION("EMPGM") = ""  Then
       If aryRecData(0,0) = 1 Then
          strChk = "checked"
       End If
    End If

    bolUpdFlds = true

    For intSub = 1 to aryRecData(0,0)
       Call Get_DB_Record(intSub)
       If CheckEmail(strEmail) = 0 Then
          If intCol = intMaxCols Then
             intCol = 0
             Response.Write "</tr><tr>" & vbCrLf
          End If
          Response.Write "<td nowrap>" & vbCrLf
          intCol = intCol + 1
          strLastEmUsed = strEmail
          strLastNameUsed = ReplChars(strEmail & " " & ReplChars(strFName & "~" & strLName," ","~"),"'","")
          If instr(SESSION("EMN"),strLastEmUsed) = 0 Then
             Response.Write "<input type='hidden' name='EMN' value='" & strLastNameUsed & "'>" & vbCrLf
          End If
          If strSecEM = "Y" Then
             strHld = ""
          Else
             strHld = MbrDetails(strFName,strLName,strEmail,strContactInfo,strCompany,strKeywords)
             strHld = "oncontextmenu=" & CHR(34) & "return ShowDetails('" & strHld & "')" & CHR(34) & " "
          End If
          Response.Write "<input type='checkbox' name='EM' value='" & strEmail & "' " & strHld & strChk & ">" & vbCrLf
          strHld = trim(strLName)
          If strHld <> "" and trim(strFName) <> "" Then
             strHld = strHld & ", " & trim(strFName)
          ElseIf strHld = "" Then
             strHld = trim(strFName)
          End If
          Response.Write "<b>" & strHld & "</b></td>" & vbCrLf
          intTotSel = intTotSel + 1
       End If
    Next
    Response.Write "</tr>" & vbCrLf

    If SESSION("VCF") = "H" Then
       Response.Write "<tr><td>&nbsp;</td></tr>" & vbCrLf
       Response.Write "<tr><td><b>Web Page Title:</b></td>" & vbCrLf
       Response.Write "<td colspan=" & intMaxCols-1 & ">" & vbCrLf
       strHld = GetCookie("BLDEMAIL","WEBTITLE")
       Response.Write "<input type='Text' name='WEBTITLE' size='50' maxlength='50' Value='" & strHld & "'>" & vbCrLf
       Response.Write "</td></tr>" & vbCrLf
    End If

    Response.Write "<tr><td colspan=" & intMaxCols & "><hr width='100%' size='5' noshade></td></tr>" & vbCrLf
    Response.Write "<tr><td colspan=" & intMaxCols & " align='center'>" & vbCrLf
    If intTotSel > 1 Then
       Response.Write "<input type='submit' name='ALL' value='Select All'>&nbsp;" & vbCrLf
       Response.Write "<input type='submit' name='NONE' value='Un-Select All'>&nbsp;" & vbCrLf
    End If
    If SESSION("VCF") = "Y" Then
       strChk = "Create VCF Records"
    ElseIf SESSION("VCF") = "H" Then
       strChk = "Create Web Page"
    Else
       strChk = "Assign TO & Build"
       IF intTotSel = 0 or (intTotSel = 1 and SESSION("EMPGM") = "") Then
          Response.Write "<input type='text' value='CC Assigned' size=11>&nbsp;" & vbCrLf
          Response.Write "<input type='text' value='BCc Assigned' size=12>&nbsp;" & vbCrLf
       Else
          Response.Write "<input type='submit' name='CC' value='Assign CC'>&nbsp;" & vbCrLf
          Response.Write "<input type='submit' name='BCC' value='Assign BCc'>&nbsp;" & vbCrLf
       End If
    End If
    If SESSION("EMPGM") <> "" Then
       If intTotSel = 0 Then
          strChk = "Return"
       Else
         strChk = "Assign TO & Return"
       End If
       Response.Write "<input type='submit' name='RETURN' value='" & strChk & "'>&nbsp;" & vbCrLf
    ElseIf intTotSel = 0 Then
       Response.Write "<input type='text' value='" & strChk & "' size=18>" & vbCrLf
    Else
       Response.Write "<input type='submit' name='BUILD' value='" & strChk & "'>" & vbCrLf
    End If

    If SESSION("SECLEVL") < 3 Then
       If SESSION("MENUFN") <> "" Then
         Response.Write "<input type=" & CHR(34) & "button" & CHR(34) & " value=" & CHR(34) & "Menu" & CHR(34) & " onClick=" & CHR(34) & "document.location.href = '" & SESSION("MENUFN") & "'" & CHR(34) & ">" & vbCrLf
       End If
       If strDirHelpFn <> "" Then
         Response.Write "<input type=" & CHR(34) & "button" & CHR(34) & " value=" & CHR(34) & "Help" & CHR(34) & " onClick=" & CHR(34) & "window.open('" & strDirHelpFn & "')" & CHR(34) & ">" & vbCrLf
       End If
    End If

    Response.Write "</form></td></tr>" & vbCrLf
    strMaint = "Y"
    call Wrapup_Web_Page

    Response.Write "<table align='center'>" & vbCrLf
    If intTotSel > 0 and strSecEM <> "Y" Then
       Response.Write "<tr><td colspan='2' align='center'><font size='-1' color='red'>Right click on a checkbox to view member's details.</font></td></tr>" & vbCrLf
    End If
    If SESSION("VCF") = "H" Then
       Response.Write "<tr><td colspan='2' align='center'><font size='-1' color='red'>Popup windows must be enabled to create web page.</font></td></tr>" & vbCrLf
    End If
    If SESSION("VCF") = "" Then
       If SESSION("TO") <> "" or SESSION("CC") <> "" or SESSION("BCC") <> "" Then
          Response.Write "<tr><td nowrap align='center' colspan='2' font=+1><br><br><b>Current Recipients</b><br></td><td></td></tr> " & vbCrLf
          Call Display_Current_Selections("TO")
          Call Display_Current_Selections("CC")
          Call Display_Current_Selections("BCC")
       End if
    End if
    Response.Write "</table>" & vbCrLf

End SUB

SUB Display_Current_Selections(TYP)

    DIM strHld, intLoc, strAddr, strName

    If SESSION(TYP) <> "" Then
       Response.Write "<tr><td valign='top'>" & vbCrLf
       Response.Write "<b>" & TYP & ":</b>&nbsp;&nbsp;" & vbCrLf
       Response.Write "</td>" & vbCrLf
       Response.Write "<td width=700>" & vbCrLf
       Response.Write SESSION(TYP) & vbCrLf
       Response.Write "</td></tr>" & vbCrLf
    End If

END SUB

SUB Verify_Address_Variables(TYP)

    DIM strHld, intLoc, strAddr, strName

    If SESSION(TYP) <> "" Then
       strHld = trim(ReplChars(SESSION(TYP),",",";"))
       SESSION(TYP) = ""
       do while strHld <> ""
          intLoc = instr(strHld,"{")
          If intLoc = 0 Then
             intLoc = len(strHld)+1
          Else
             strName = trim(left(strHld,intLoc-1))
             strHld = mid(strHld,intLoc+1)
             intLoc = instr(strHld,"}")
             If intLoc = 0 Then
                intLoc = len(strHld)+1
             End If
          End If
          strAddr = trim(left(strHld,intLoc-1))
          If SESSION(TYP) = "" Then
             SESSION(TYP) = strAddr
          Else
             SESSION(TYP) = SESSION(TYP) & ", " & strAddr
          End If
          If instr(SESSION("EMN"),strAddr) = 0 and trim(strName) <> "" Then
             strName = ReplChars(strName," ","~")
             If SESSION("EMN") = "" Then
                SESSION("EMN") = strAddr & " " & strName & ", "
             Else
                SESSION("EMN") = SESSION("EMN") & ", " & strAddr & " " & strName
             End if
          End If
          strHld = mid(strHld,intLoc+1)
       loop
    End If

END SUB

SUB Build_VCF_HTML_Entries

    DIM intCntr, intTotal, strTitle

    If SESSION("VCF") = "H" Then
       strTitle = trim(Request.Form("WEBTITLE"))
       If strTitle = "" Then
          strTitle = GetCookie("BLDEMAIL","WEBTITLE")
       End If
       Call AddCookie("BLDEMAIL","WEBTITLE",strTitle,365)
       If not bolNPOption Then
          SESSION("V-TO") = strToAddr
          Response.Redirect "Directory_Web_Page_Build.asp"
       Else
          strToAddr = SESSION("V-TO")
          SESSION("V-TO") = ""
          intCols = 0
          Response.Write "<html>" & vbCrLf
          Response.Write "<head>" & vbCrLf
          Response.Write "<title>" & strTitle & "</title>" & vbCrLf
          CALL Setup_ShowDetails_Script
          Response.Write "</head>" & vbCrLf
          Response.Write "<body link='Blue' vlink='Blue' alink='Blue'>" & vbCrLf
          Response.Write "<div align='center'><font size=+1><b>" & strTitle & "</b></font></div>" & vbCrLf
          Response.Write "<table border='1' cellpadding='4' cellspacing='0' align='center'>" & vbCrLf
          Response.Write "<tr>" & vbCrLf
       End If
    Else
       Response.Write "<b><font color='red'>" & vbCrLf
       If bolRexx Then
          Response.Write "<ol type='1'>" & vbCrLf
          Response.Write "<li>Highlight all of the text from the first BEGIN: to the last END: and copy into the clipboard</li>" & vbCrLf
          Response.Write "<li>Using Notepad, open the file C:\TEMP\VCF.TXT</li>" & vbCrLf
          Response.Write "<li>Paste the clipboard text into the opened file and close</li>" & vbCrLf
          Response.Write "<li>Run SPLIT_VCF_FILE to split VCF.TXT into individual VCF Import Files, or</li>" & vbCrLf
          Response.Write "<li>Run VCF_2_COMMA_DEL to create the comma delimited file VCF.CSV</li>" & vbCrLf
          Response.Write "<li>Import each VCF file or the VCF.CSV file into the Address Book</li>" & vbCrLf
       Else
          Response.Write "For each list entry:" & vbCrLf
          Response.Write "<ol type='1'>" & vbCrLf
          Response.Write "<li>Highlight the BEGIN: to END: text</li>" & vbCrLf
          Response.Write "<li>Copy the text into the clipboard</li>" & vbCrLf
          Response.Write "<li>Using Notepad, open the file C:\TEMP\FirstName_LastName.vcf</li>" & vbCrLf
          Response.Write "<li>Paste the clipboard text into the opened file and close</li>" & vbCrLf
          Response.Write "<li>Repeat steps 1-4 for the remaining entries</li>" & vbCrLf
          Response.Write "<li>Import each VCF file in C:\TEMP into the Address Book</li>" & vbCrLf
       End If
          Response.Write "</ol></b></font><br>" & vbCrLf
    End If

    If trim(strToAddr) <> "" Then

       strToAddr = trimspaces(strToAddr,1)
       intTotal =  CountChars(strToAddr," ") + 1
       strToAddr = " " & strToAddr & " "

       bolBuildAryOnly = true
       Call Directory_Selection_List("")

       bolUpdFlds = true

       For intSub = 1 to aryRecData(0,0)
          Call Get_DB_Record(intSub)
          intCntr = 0
          If instr(strToAddr," " & strEmail & " ") > 0 Then
             If SESSION("VCF") = "H" Then
                CALL Build_HTML_Entry(strFName,strLName,strEmail,strContactInfo,strCompany,strKeywords)
             Else
                CALL Display_VCF_Entry(strFName,strLName,strEmail,strContactInfo,strCompany)
             End if
             intCntr = intCntr + 1
          End If
          If intCntr >= intTotal Then
             EXIT FOR
          End If
       Next
    End If

    If SESSION("VCF") = "H" Then
       If bolNPOption Then
          Response.Write "</tr></table>" & vbCrLf
          Response.Write "<div align='center'><font size='-1' color='red'><br>" & vbCrLf
          Response.Write "<u>Left</u> click on a name to open a new email message<br>" & vbCrLf
          Response.Write "<u>Right</u> click on a name to view the directory details" & vbCrLf
          Response.Write "</font></div>" & vbCrLf
          Response.Write "</body></html>" & vbCrLf
       Else
          intTblCols = intTblCols + 1
          Response.Write "</tr><tr><td align='center' colspan=" & intTblCols & ">" & vbCrLf
          Response.Write "<input type='button' value='Create Offline Version' onclick=" & CHR(34) & "window.open('build_email_list.asp?NP=Y')" & CHR(34) & ">" & vbCrLf
          Response.Write "</td></tr>" & vbCrLf
          strMaint = "Y"
          call Wrapup_Web_Page
       End If
    End If

END SUB

SUB Display_VCF_Entry(FName,LName,eMail,ConInfo,Company)

    DIM strConInfo, strPh, strTy, intSub, strWord

    Response.Write "BEGIN:VCARD<br>" & vbCrLf
    Response.Write "VERSION:2.1<br>" & vbCrLf

    Response.Write "N:" & LName & ";" & FName & "<br>" & vbCrLf
    Response.Write "FN:" & LName & ", " & FName & "<br>" & vbCrLf
    Response.Write "ORG:" & Company & "<br>" & vbCrLf

    If right(ConInfo,1) <> ")" Then
       Response.Write "TEL;WORK:" & ConInfo & "<br>" & vbCrLf
    Else
       strConInfo = ReplChars(ConInfo,"("," ")
       strConInfo = ReplChars(strConInfo,")"," ")
       strConInfo = ReplChars(strConInfo,"-"," ")
       strConInfo = ReplChars(strConInfo,"."," ")
       strConInfo = ReplChars(strConInfo,CHR(13)," ")

       bolDoPhrases = false
       intSub = 1
       strWord = GetWords(strConInfo,intSub,1)

       Do While strWord <> ""
          strPh = "(" & strWord & ") "
          intSub = intSub + 1
          strWord = GetWords(strConInfo,intSub,1)
          strPh = strPh & strWord & "-"
          intSub = intSub + 1
          strWord = GetWords(strConInfo,intSub,1)
          strPh = strPh & strWord
          intSub = intSub + 1
          strWord = GetWords(strConInfo,intSub,1)
          If strWord = "H" Then
             strTy = "HOME"
          ElseIf strWord = "M" or strWord = "C" Then
             strTy = "CELL"
          ElseIf strWord = "F" Then
             strTy = "FAX"
          Else
             strTy = "WORK"
          End If
          Response.Write "TEL;" & strTy & ":" & strPh & "<br>" & vbCrLf
          intSub = intSub + 1
          strWord = GetWords(strConInfo,intSub,1)
       Loop
    End If
    Response.Write "EMAIL:" & eMail & "<br>" & vbCrLf
    Response.Write "END:VCARD<br><br>" & vbCrLf
END SUB

SUB Build_HTML_Entry(FName,LName,eMail,ConInfo,Company,KeyWords)

    DIM strName, strJS

    If trim(eMail) = "" Then
       EXIT SUB
    End If

    strJS = MbrDetails(FName,LName,eMail,ConInfo,Company,KeyWords)
    strJS = "oncontextmenu=" & CHR(34) & "return ShowDetails('" & strJS & "')" & CHR(34) & " "

    strName = LName
    If trim(FName) <> "" Then
       If strName <> "" Then
          strName = strName & ", " & FName
       Else
          strName = FName
       End If
    End If

    intTblCols = 6
    If intCols > intTblCols Then
       Response.Write "</tr><tr>" & vbCrLf
       intCols = 0
    End If
    Response.Write "<td nowrap><a href=" & CHR(34) & "mailto:" & FName & " " & LName  & "<" & eMail & ">" & CHR(34) & " " & strJS & "> <font size='-1'>" & strName & "</font></a></td>" & vbCrLf
    intCols = intCols + 1

END SUB

FUNCTION MbrDetails(FName,LName,eMail,ConInfo,Company,KeyWords)

    DIM strPH, strTemp, strWord, intSub

    MbrDetails = FName & " " & LName & "\n" &eMail
    strPH = trim(ConInfo)
    If left(strPH,1) = "(" Then
       MbrDetails =MbrDetails & "\n" & ReplLineBreaksJS(strPH)
    End If
    If trim(Company) <> "" and instr(ucase(SESSION("KW")),ucase(Company)) = 0 Then
       MbrDetails =MbrDetails & "\n" & Company
    End If

    bolDoPhrases = true

    Keywords = RemoveLineBreaks(Keywords)
    intSub = 1
    strTemp = GetWords(Keywords,intSub,1)
    intSub = intSub + intRetWords
    strWord = GetWords(Keywords,intSub,1)

    do while strWord <> ""
       strTemp = strTemp & ", " & strWord
       intSub = intSub + intRetWords
       strWord = GetWords(Keywords,intSub,1)
    Loop

    If strTemp <> "" Then
       MbrDetails =MbrDetails & "\n" & strTemp
    End If

    MbrDetails = ReplChars(MbrDetails,"'","\'")

END FUNCTION

FUNCTION CheckEmail(EM)

   If SESSION("VCF") <> "" Then
      CheckEmail = 0
   Else
      CheckEmail = instr(SESSION("TO"),EM) + instr(SESSION("CC"),EM) + instr(SESSION("BCC"),EM)
   End If

END FUNCTION

SUB Get_Security_Members

    DIM intSub

    bolBldSecurAryOnly = true
    Call Security_Selection_List("")

    For intSub = 1 to arySecurity(0,0)
       aryRecData(intSub,1) = arySecurity(intSub,1)   'ID
       aryRecData(intSub,2) = arySecurity(intSub,10)  'First Name
       aryRecData(intSub,3) = arySecurity(intSub,11) 'Last Name
       aryRecData(intSub,4) = arySecurity(intSub,4)   'Email
       aryRecData(intSub,6) = arySecurity(intSub,7)  'Security Groups
    Next
    aryRecData(0,0) = arySecurity(0,0)
END SUB
%>


<!--#INCLUDE FILE="../dir/cookie_maint.inc" -->

<%

   DIM arySrchOpts(5,2), aryStatusOpts(2),arySpecOpt(5), bolNoAdvSrch, intSrchRows, bolBottomButton, bolSrchAryOnly
   DIM bolDefSpecOpt

SUB Insert_Search_Prompt(PGM)

   DIM intCols, intLoc, strHld, intSub, intPos

   If intSrchRows = "" or not isNumeric(intSrchRows) or intSrchRows < 1 Then
      intSrchRows = 1
   End If

   Response.Write "<table>" & vbCrLf
   Response.Write "<tr>" & vbCrLf
   Response.Write "<form action='" & PGM & "' method='post' name='Form' id='Form'>" & vbCrLf
   Response.Write "<input type='hidden' name='SEARCH' value='Y'>" & vbCrLf
   If SESSION("AS") <> "Y" Then
      Response.Write "<td nowrap>" & vbCrLf
      Response.Write "<a href='" & strSrchHelpFn & "' onmouseover=" & CHR(34) & "window.status='Search Documentation';return true" & CHR(34) & "><font face='Times New Roman, Times, serif' color='#666633' size=-2>Search Help</font></a>" & vbCrLf
      Response.Write "</td><td align='right'>" & vbCrLf
      If not bolNoAdvSrch Then
         Response.Write "<a href='" & PGM & "?AS=Y&P=-' onmouseover=" & CHR(34) & "window.status='Advance Search Option';return true" & CHR(34) &"><font face='Times New Roman, Times, serif' color='#666633' size=-2>Advanced Search</font></a>" & vbCrLf
      End If
      Response.Write "</td><td&nbsp;</td>" & vbCrLf
   Else
      Response.Write "<td nowrap>" & vbCrLf
      Response.Write "<font face='Times New Roman, Times, serif' color='#666633' size=2>" & vbCrLf
      If arySrchOpts(0,0) = "" Then
         Response.Write "&nbsp;</td>" & vbCrLf
      Else
         Response.Write "Search Area:<br>" & vbCrLf
         For intLoc = 1 to len(arySrchOpts(0,0))
            If intLoc = 1 Then
               If trim(SESSION("SRCHAREA")) = "" Then
                  strHld  = " checked"
               Else
                  strHld = ""
               End If
               Response.Write "<input type='radio' value='' name='SA'" & strHld & ">All" & vbCrLf
            End If
            If mid(arySrchOpts(0,0),intLoc,1) = SESSION("SRCHAREA") Then
               strHld  = " checked"
            Else
               strHld = ""
            End If
            If arySrchOpts(intLoc,1) <> "" Then
               intPos = instr(arySrchOpts(intLoc,1),"#PGM#")
               If intPos > 0 Then
                  arySrchOpts(intLoc,1) = left(arySrchOpts(intLoc,1),intPos-1) & PGM & mid(arySrchOpts(intLoc,1),intPos+5)
               End if
               strHld = strHld & arySrchOpts(intLoc,1)
            End If
            Response.Write "<input type='radio' value='" & mid(arySrchOpts(0,0),intLoc,1) & "' name='SA'" & strHld & ">" & arySrchOpts(intLoc,0)  & vbCrLf
         Next
         Response.Write "</font>" & vbCrLf
         Response.Write "</td>" & vbCrLf
      End If
      If aryStatusOpts(0) = "" Then
         Response.Write "<td>&nbsp;</td>" & vbCrLf
      Else
         Response.Write "<td nowrap>" & vbCrLf
         Response.Write "<font face='Times New Roman, Times, serif' color='#666633' size=2>" & vbCrLf
         Response.Write "Search Status:<br>" & vbCrLf
         For intLoc = 1 to len(aryStatusOpts(0))
            If mid(aryStatusOpts(0),intLoc,1) = SESSION("SRCHTYPE") Then
               strHld  = " checked"
            Else
               strHld = ""
            End If
            Response.Write "<input type='radio' value='" & mid(aryStatusOpts(0),intLoc,1) & "' name='TY'" & strHld & ">" & aryStatusOpts(intLoc)  & vbCrLf
            If intLoc = len(aryStatusOpts(0))  Then
               If trim(SESSION("SRCHTYPE")) = "" Then
                  SESSION("SRCHTYPE") = mid(aryStatusOpts(0),intLoc,1)
                  strHld  = " checked"
               Else
                  strHld = ""
               End If
               Response.Write "<input type='radio' value='' name='TY'" & strHld & ">All" & vbCrLf
            End If
         Next
         Response.Write "</font>" & vbCrLf
         Response.Write "</td>" & vbCrLf
      End If
      Response.Write "<td>&nbsp;</td>" & vbCrLf
   End If
   Response.Write "</tr>" & vbCrLf
   Response.Write "</td><td nowrap colspan=2>" & vbCrLf
   If SESSION("AS") = "Y" Then
      intCols = 45
   Else
      intCols = 30
   End If
   If intSrchRows = 1 Then
      Response.Write "<input type='text' name='Keywords' size=" & intCols & " value='" & SESSION("SRCHKW") & "'>" & vbCrLf
   Else
      Response.Write "<textarea name='Keywords' cols='70' rows='" & intSrchRows & "' wrap='YES'>" & SESSION("SRCHKW") & "</textarea>" & vbCrLf
   End If
   Response.Write "</td><td>" & vbCrLf
   If not bolBottomButton Then
      Response.Write "<input type='image' src='/" & Application("PGMDIR") & "/graphics/searchbut.gif' alt='Search'>" & vbCrLf
   Else
      Response.Write "&nbsp;" & vbCrLf
   End If
   Response.Write "</td></tr>" & vbCrLf
   Response.Write "<tr><td nowrap valign='top'>" & vbCrLf

   If SESSION("AS") = "Y" Then
      Response.Write "<a href='" & strSrchHelpFn & "' onmouseover=" & CHR(34) & "window.status='Book Search Documentation';return true" & CHR(34) & "><font face='Times New Roman, Times, serif' color='#666633' size=-2>Search Help</font></a>" & vbCrLf
   Else
      Response.Write "&nbsp;" & vbCrLf
   End If
   Response.Write "</td><td nowrap align='right'>" & vbCrLf
   If arySpecOpt(0) > 0 Then
      If not isnumeric(SESSION("SPOPER")) or SESSION("SPOPER") = "" Then
         SESSION("SPOPER") = arySpecOpt(0)
      End If
      Response.Write "<font face='Times New Roman, Times, serif' color='#666633' size=2>" & vbCrLf
      For intSub = 1 to arySpecOpt(0)
         If bolDefSpecOpt Then
            Response.write "<input type='hidden' name='SpOpt' value='1'>" & vbCrLf
            EXIT FOR
         Else
            If int(SESSION("SPOPER")) = intSub Then
               strHld = " checked"
            Else
               strHld = ""
            End If
            If arySpecOpt(0) = 1 Then
               Response.write "<input type='Checkbox' name='SpOpt' value='1'" & strHld & ">" & vbCrLf
            Else
               Response.write "<input type='Radio' name='SpOpt' value='" & intSub & "'" & strHld & ">" & vbCrLf
            End If
            Response.Write "<b>" & arySpecOpt(intSub) & "</b>" & vbCrLf
         End If
      Next
      Response.Write "</font>&nbsp;&nbsp;&nbsp;" & vbCrLf
   End If
   If bolBottomButton Then
      Response.Write "<input type='image' src='/" & Application("PGMDIR") & "/graphics/searchbut.gif' alt='Search'>" & vbCrLf
   Else
      Response.Write "&nbsp;" & vbCrLf
   End If
   Response.Write "</td><td></td>" & vbCrLf
   Response.Write "</form></tr></table>" & vbCrLf
END SUB

SUB Check_Search_Options

    If Request.QueryString("AS") <> "" Then
       SESSION("AS") = Request.QueryString("AS")
    End If

   If Request.QueryString("RKW") <> "" Then
      strTempKW = Request.QueryString("RKW")
      If ucase(strTempKW) = "*ALL*" Then
         strTempKW = ""
      End If
      SESSION("SRCHAREA") = "C"
      Session.Contents.Remove("SRCHSQL")
   End If

   If Request.Form("SEARCH") = "Y" Then
      SESSION("SRCHAREA") = Request.Form("SA")
      SESSION("SRCHTYPE") = Request.Form("TY")
      SESSION("SRCHKW") = Request.Form("Keywords")
      SESSION("SPOPER") = Request.Form("SpOpt")
      Session.Contents.Remove("SRCHSQL")
   End IF

   If SESSION("SRCHAREA")  = "C" Then
      bolGenCatOnly = true
   Else
      bolGenTxtOnly = true
   End If

END SUB

SUB Get_Search_Cookies(OPT)

   DIM strCookieName

   strCookieName = OPT&"SRCH"

   If Request.Form("SEARCH") = "Y" Then
      EXIT SUB
   End If

   If SESSION("SRCHAREA") = "" Then
      SESSION("SRCHAREA") = GetCookie(strCookieName,"SRCHAREA")
   End If
   If SESSION("SRCHTYPE") = "" Then
      SESSION("SRCHTYPE") = GetCookie(strCookieName,"SRCHTYPE")
   End If
   If SESSION("SRCHKW") = "" Then
      SESSION("SRCHKW") = GetCookie(strCookieName,"SRCHKW")
   End If
   If SESSION("SPOPER") = "" Then
      SESSION("SPOPER") = GetCookie(strCookieName,"SPOPER")
   End If

END SUB

SUB Update_Search_Cookies(OPT)

   DIM strCookieName

   strCookieName = OPT&"SRCH"

   Call AddCookie(strCookieName,"SRCHAREA",SESSION("SRCHAREA"),365)
   Call AddCookie(strCookieName,"SRCHTYPE",SESSION("SRCHTYPE"),365)
   Call AddCookie(strCookieName,"SRCHKW",SESSION("SRCHKW"),365)
   Call AddCookie(strCookieName,"SPOPER",SESSION("SPOPER"),365)

END SUB

FUNCTION CheckSearch

   If Request.Form("SEARCH") = "Y" or SESSION("SRCHKW") <> "" or SESSION("CAT") <> "" or Request.QueryString("RKW") <> "" Then
      CheckSearch = true
   Else
      CheckSearch = false
   End If

END FUNCTION


%>


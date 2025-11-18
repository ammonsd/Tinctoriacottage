<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/maint_webpage.inc" -->
<!--#INCLUDE FILE="../dir/directory_select.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->
<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/directory_navigation_details.inc" -->

<%
    DIM intSub, strHdr, intNbrCols, bolPhone, bolNotes, bolCompany, bolAll, strWord, strHld, intCntr
    DIM bolGetSrchCriteria, bolSecMaintAccess, strGoTo, strFullName, bolFnLn

    strGoTo = "directory_list.asp"

    If Request.QueryString("DKW") <> "Y" Then
       If Request.QueryString("DT") <> "" or Request.QueryString("EL") = "Y"  Then
          SESSION("DT") = Request.QueryString("DT")
          SESSION("NA") = Request.QueryString("NA") 'No Address in Display
          If Request.QueryString("DEF") = "" Then
             SESSION("DEF") = "D"
          Else
             SESSION("DEF") = Request.QueryString("DEF")
          End If
          SESSION("SPO") = Request.QueryString("SPO")
          Session.Contents.Remove("RKW")
          Session.Contents.Remove("SkEmEqCo")
       End If
    End If

    intNbrCols = 2

    If instr(SESSION("DT"),"P") > 0 Then
       bolPhone = true 'Contact Info
       intNbrCols = intNbrCols + 1
    End If
    If instr(SESSION("DT"),"N") > 0 Then
       bolNotes = true 'Notes
       intNbrCols = intNbrCols + 1
    End If
    If instr(SESSION("DT"),"C") > 0 Then
       bolCompany = true 'Company
       intNbrCols = intNbrCols + 1
    End If

    If not bolPhone and not bolNotes and not bolCompany Then
       bolAll = true
       intNbrCols = intNbrCols + 3
    End If

    If Request.QueryString("E2C") <> "" Then
       SESSION("SkEmEqCo") = Request.QueryString("E2C")
    End If
    If Request.QueryString("SN") <> ""  Then
       SESSION("SN") = Request.QueryString("SN")
    End If
    If Request.QueryString("SE") <> ""  Then
       SESSION("SE") = Request.QueryString("SE")
    End If

    call Database_Setup
    Call System_Setup(strGoTo)

    If Application("EnableSrch") = "Y" Then
       If Request.Form("SEARCH") <> "Y" Then
          SESSION("AS") = "Y" 'Turn on Advance Search Options
          bolSrchAryOnly = true
          Call Check_Search_Criteria("Y")
          SESSION("BC") = ""
          call Setup_Web_Page("Enter Search Criteria",2)
          Response.Write "<tr><td>" & vbCrLf
          intSrchRows = 10
          bolBottomButton = true
          Call Insert_Search_Prompt("directory_list.asp")
          Response.Write "</td></tr>" & vbCrLf
          Call Insert_Navigation_Details("S")
          strMaint = "Y"
          call Wrapup_Web_Page
          Response.End
       End If
    End If

    Call Check_Search_Criteria("Y")

    If Request.QueryString("G") <> ""  Then
       Call Process_External_Keywords(Request.QueryString("G"))
       SESSION("SO") = Request.QueryString("SO")
       If Request.QueryString("DKW") = "Y" Then
          strGoTo=strGoTo&"?DKW=Y"
       ElseIf Request.QueryString("NS") = "Y" Then
          strGoTo=strGoTo&"?NS=Y"
       End If
       Response.Redirect(strGoTo)
    End If

    If Request.Form("SpOpt") <> "" Then
       SESSION("SPOPER") = Request.Form("SpOpt")
       SESSION("SRCHKW") = Request.Form("Keywords")
    End If

    If (instr("124",Request.Form("SpOpt")) > 0 and Request.Form("SpOpt") <> "") or Request.QueryString("EL") = "Y" or Request.QueryString("NS") = "Y" or Request.QueryString("DKW") = "Y" Then
       If Request.QueryString("EL") = "Y" Then
          SESSION("emKW") = Request.QueryString("RKW")
       Else
          SESSION("SRCHKW") = Request.Form("Keywords")
       End If
       SESSION("EMN") = ""
       If Request.QueryString("DKW") = "Y" Then
          Response.Redirect(strLstKwFn & Application("DIRTBL") & "&NC=Y")
       Else
          If Request.Form("SpOpt") = "2" or SESSION("DEF") = "W" Then
             strHld = "&V=H"
          ElseIf Request.Form("SpOpt") = "4" or SESSION("DEF") = "V" Then
             strHld = "&V=Y"
          End If
          SESSION("DIRADDR") = "Y"
          Response.Redirect("build_email_list.asp?SO=Y" & strHld)
       End If
    Else
       Session.Contents.Remove("emKW")
    End If

    If SESSION("RQL") = "" Then
       SESSION("RQL") = 3
    End If

    Call Logon_Check(GetCurPath("")&strGoTo,SESSION("RQL"),strLogonGrp)

    If SESSION("KW") <> "" Then
       call Read_Text_File(SESSION("USERID")&"_DirEmDefaults.cfg",strFileLoc)
       bolDoPhrases = true
       strHld = ucase(GetWords(SESSION("KW"),1,1))
       For intSub = 1 to aryTxtRecs(0)
          strWord = ucase(GetWords(aryTxtRecs(intSub),1,1))
          If strWord = strHld Then
             intCntr = intRetWords + 1
             SESSION("SE") = GetWords(aryTxtRecs(intSub),intCntr,1)
             intCntr = intCntr + 1
             SESSION("SN") = GetWords(aryTxtRecs(intSub),intCntr,99)
             EXIT FOR
          End If
       Next
    End If

    bolBuildAryOnly = true

    If SESSION("NA") = "Y" Then
       intNbrCols = IntNbrCols - 1
    End If

    intCS = 1
    intCP = 2

    If SESSION("SO") = "C" Then
       bolCoSort = true
    Else
       bolCoSort = false
    End If

    bolLastNameFirst = true

    If bolAll Then
       strSpecSecGrp = "SecMaint"
       bolDisplayMsg = false
       Call Verify_Group_Access("N/R","",5)
       If not bolSecError Then
          bolSecMaintAccess = true
          intNbrCols = intNbrCols + 1
       End If
    End If

    strHld = SESSION("GROUP")
    If strHld <> "" Then
       if right(ucase(strHld),1) <> "S" Then
          strHld = strHld & "'s"
       End If
    End If
    call Setup_Web_Page(strHld&" Directory List",intNbrCols)
    Call Directory_Selection_List("")

    strHdr= "<tr><td><b>Name</b></td>"
    If bolPhone or bolAll Then
       strHdr = strHdr  & "<td><b>Phone #</b></td>"
    End If
    If SESSION("NA")<> "Y" Then
       strHdr= strHdr & "<td><b>Email Address</b></td>"
    End If
    If bolCompany or bolAll Then
       strHdr = strHdr & "<td><b>Company</b></td>"
    End If
    If bolNotes or bolAll Then
       strHdr = strHdr & "<td>&nbsp;&nbsp;<b>Notes</b></td>"
    End If
    If bolAll and bolSecMaintAccess Then
       strHdr = strHdr & "<td>&nbsp;&nbsp;<b>Security Groups</b></td>"
    End If

    strHdr = strHdr & "</tr>"

    Response.Write strHdr & vbCrLf
    Response.Write "<tr><td colspan='" & intNbrCols & "'><hr align='left' size='5' noshade></td></tr>" & vbCrLf

    intCntr = 0

    bolUpdFlds = true

    For intSub = 1 to aryRecData(0,0)
       Call Get_DB_Record(intSub)
       If SESSION("SkEmEqCo") = "Y" Then
          'Do not include company if it is the same as the email address AND is part of the Company search critera
          'Show only company names that are not part of the group
          strHld = RemoveSpaces(strCompany)
          If instr(ucase(strEmail),ucase(strHld)) > 0 and instr(ucase(SESSION("KW")),ucase(strCompany)) > 0 Then
             strCompany = ""
          End If
       End If
       If (bolPhone and SESSION("NA") = "Y" ) and (trim(strContactInfo) = "" or isnull(strContactInfo)) Then
          'Skip Entry
       Else
          intCntr = intCntr + 1
          strKeyWords = ReplChars(strKeyWords,CHR(13),"<br>")
          strContactInfo = ReplChars(strContactInfo,CHR(13),"<br>")
          If bolFnLn Then
             strHld = strFName
             strFName = strLName
             strLName = strHld
          End If
          strFullName = strLName
          If strFullName <> "" and strFName <> "" Then
             If bolFnLn Then
                strFullName = strFullName & " "
             Else
                strFullName = strFullName & ", "
             End If
          End If
          strFullName = strFullName & strFName
          Response.Write "<tr>" & vbCrLf
          Response.Write "<td nowrap>" & strFullName & "</td>" & vbCrLf
          If bolPhone or bolAll Then
             Response.Write "<td nowrap>" & strContactInfo & "</td>" & vbCrLf
          End If
          If SESSION("NA") <> "Y" Then
             Response.Write "<td nowrap>" & strEmail & "</td>" & vbCrLf
          End If
          If bolCompany or bolAll Then
             Response.Write "<td nowrap>" & strCompany & "</td>" & vbCrLf
          End If
          If bolNotes or bolAll Then
             Response.Write "<td nowrap>&nbsp;&nbsp;" & strKeyWords & "</td>" & vbCrLf
          End If
          If bolAll and bolSecMaintAccess Then
             Response.Write "<td nowrap>&nbsp;&nbsp;" & strSecGrps & "</td>" & vbCrLf
          End If
          Response.Write "</tr>" & vbCrLf
       End If
    Next

    Response.Write "<tr><td colspan='" & intNbrCols & "'><hr align='left' size='5' noshade></td></tr>" & vbCrLf
    If intCntr > 5 Then
       Response.Write "<tr><td align='right' colspan='" & intNbrCols & "'><font size=-2>" & intCntr &"</font></td></tr>" & vbCrLf
    End If

    call Wrapup_Web_Page
%>

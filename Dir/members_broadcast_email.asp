<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/maint_webpage.inc" -->
<!--#INCLUDE FILE="../dir/member_select.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->
<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/member_navigation_details.inc" -->

<%
    DIM intSub, strHld, bolExpFnd, bolDisAll, strAddress

    If Request.Form("PROCESS") <> "" Then
       If Request.Form("PROCESS") = "Send Email to Each" Then
          SESSION("IM") = "Y"
       Else
          SESSION("IM") = "N"
       End If
       Response.Redirect("email.asp")
    End If

    Call System_Setup("NONE")
    strTypeSel = "A"
    bolBuildAryOnly = true
    intCS = 3

    Call Logon_Check(GetCurPath("")&"members_broadcast_email.asp",3,strLogonGrp)

    If Application("EnableSrch") = "Y" Then
       If Request.Form("SEARCH") <> "Y" Then
          SESSION("AS") = "Y" 'Turn on Advance Search Options
          bolSrchAryOnly = true
          Call Check_Search_Criteria
          call Setup_Web_Page("Enter Search Criteria",1)
          Response.Write "<tr><td>" & vbCrLf
          intSrchRows = 5
          bolBottomButton = true
          Call Insert_Search_Prompt("members_broadcast_email.asp")
          Response.Write "</td></tr>" & vbCrLf
          Call Insert_Navigation_Details("S")
          strMaint = "Y"
          call Wrapup_Web_Page
          Response.End
       End If
    End If

    Call Check_Search_Criteria

    If Request.QueryString("IM") <> "" Then
       SESSION("IM") = Request.QueryString("IM")
    End If

    call Setup_Web_Page(strCompany & " eMail Listing",intCS)

    Call Get_DB_Entries

    strHld = "<tr><td align='center'><b>Name</b></td></td><td align='center'><b>Email</b></td>"
    Response.Write strHld & "</tr>" & vbCrLf
    Response.Write "<tr><td colspan='" & intCS & "'><hr align='left' size='5' noshade></td></tr>" & vbCrLf

    bolUpdFlds = true
    SESSION("ML") = ""

    For intSub = 1 to aryRecData(0,0)
       call Get_DB_Record(intSub)
       Response.Write "<tr>" & vbCrLf
       Response.Write "<td valign='top'>" & trim(strFName & " " & strLName) & "</td>" & vbCrLf
       Response.Write "<td valign='top'>" & strEmail & "</td>" & vbCrLf
       Response.Write "</tr>" & vbCrLf
       If SESSION("ML") <> "" Then
          SESSION("ML") = SESSION("ML") & ", "
       End If
       SESSION("ML") = SESSION("ML") & trim(strFName & " " & strLName) & " {" & strEmail & "}"
    Next

    Response.Write "<tr><td><p>&nbsp;</p></td></tr>" & vbCrLf
    If aryRecData(0,0) > 0 Then
       Response.Write "<tr><td colspan=2 align='center'>" & vbCrLf
       Response.Write "<form Action='members_broadcast_email.asp' Method='post'>" & vbCrLf
       Response.Write "<input type='Submit' Name='PROCESS' value='Bcc ALL with One Email'>" & vbCrLf
       Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
       Response.Write "<input type='Submit' Name='PROCESS' value='Send Email to Each Member'>" & vbCrLf
       Response.Write "</form>" & vbCrLf
       Response.Write "</td></tr>" & vbCrLf
    End If

    bolNoLinks = true
    call Wrapup_Web_Page

%>

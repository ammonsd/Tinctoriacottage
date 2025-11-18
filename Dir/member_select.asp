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

    Call System_Setup("NONE")
    Call Logon_Check(GetCurPath("")&"member_select.asp",3,strLogonGrp)

    bolGoToUrl = true
    bolAutoSel  = true

    If Application("EnableSrch") = "Y" Then
       SESSION("AS") = "Y" 'Turn on Advance Search Options
    End If

    Call Check_Search_Criteria

    strMaint = "Y"

    call Setup_Web_Page("Select Member Entry",1)

    Response.Write "<tr><td nowrap align='center'>" & vbCrLf
    Response.Write "<form action='member_maint.asp?TM=U' method='post'>" & vbCrLf

    strGoToPgm = "member_maint.asp?TM=U&ID="
    If Request.Form("SEARCH") = "Y" or Request.QueryString("RKW") <> "" Then
       Call Get_DB_Entries
    Else
       Response.Write "<font color='red'><br><b>Search must be Invoked to Obtain Member List</b></font>" & vbCrLf
    End If

    If aryRecData(0,0) > 0 Then
       If bolAutoSel Then
          Response.Write "<input type='submit' value='Select' onClick='goUrl(this.form.Url);return false;'>" & vbCrLf
       Else
          Response.Write "<input type='submit' value='Select'>" & vbCrLf
       End If
    End If
    Response.Write "</form>" & vbCrLf
    Response.Write "</td></tr>" & vbCrLf

    IF Application("EnableSrch") = "Y" Then
       Response.Write "<tr><td>" & vbCrLf
       Response.Write "<br><br><br><br><br>" & vbCrLf
       intSrchRows = 5
       bolBottomButton = true
       Call Insert_Search_Prompt("member_select.asp")
       Response.Write "</td></tr>" & vbCrLf
       Call Insert_Navigation_Details("S")
    End If

    call Wrapup_Web_Page

%>

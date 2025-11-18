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

    Call System_Setup("NONE")
    Call Logon_Check(GetCurPath("")&"directory_select.asp",3,strLogonGrp)

    bolGoToUrl = true
    bolAutoSel  = true
    Session.Contents.Remove("DEF")
    Session.Contents.Remove("SPO")

    If Application("EnableSrch") = "Y" Then
       SESSION("AS") = "Y" 'Turn on Advance Search Options
    End If

    Call Check_Search_Criteria("Y")

    strTypeGoto = "WF"

    call Setup_Web_Page("Select Entry",1)

    Response.Write "<tr><td nowrap align='center'>" & vbCrLf
    Response.Write "<form action='directory_maint.asp?TM=U' method='post'>" & vbCrLf

    strGoToPgm = "directory_maint.asp?TM=U&ID="
    If Request.Form("SEARCH") = "Y" or Request.QueryString("RKW") <> "" Then
       Call Get_DB_Entries
    Else
       Response.Write "<font color='red'><br><b>Search must be Invoked to Obtain Selection List</b></font>" & vbCrLf
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
       SESSION("SPO") = ""
       SESSION("DEF") = "D"
       Call Insert_Search_Prompt("directory_select.asp")
       Response.Write "</td></tr>" & vbCrLf
       Call Insert_Navigation_Details("S")
    End If

    strMaint = "Y"
    call Wrapup_Web_Page

%>

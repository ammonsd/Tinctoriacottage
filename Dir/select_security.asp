<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/security_select.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->
<!--#INCLUDE FILE="../dir/security_maint_links.inc" -->
<!--#INCLUDE FILE="../dir/logon_check.inc" -->

<%

    strSpecSecGrp = "SecMaint"
    Call System_Setup("NONE")
    Call Logon_Check(GetCurPath("")&"select_security.asp","5",strLogonGrp)

    strMaint = Request.Form("MT")
    If strMaint = "" Then
       strMaint = Request.Querystring("MT")
    End If

    strTypeMaint = Request.Form("TM")
    If strTypeMaint = "" Then
       strTypeMaint = Request.Querystring("TM")
    End If

    bolGoToUrl = true
    bolSecAutoSel  = true

    strSecGoToPgm = Request.Form("PG")
    If strSecGoToPgm = "" Then
       strSecGoToPgm = Request.Querystring("PG")
    End If

    If strSecGoToPgm = "" Then
       strSecGoToPgm="security_maint.asp?TM=U"
    End If

    If bolSecAutoSel Then
       If instr(strSecGoToPgm,"?") = 0 Then
          strSecGoToPgm = strSecGoToPgm & "?LID="
       Else
          strSecGoToPgm = strSecGoToPgm & "&LID="
       End If
    End If

    call Setup_Web_Page("Select Profile",1)

    Response.Write "<tr><td nowrap>" & vbCrLf
    Response.Write "<form action='" & strSecGoToPgm & "' method='post'>" & vbCrLf
    Response.Write "<input type='hidden' name='TM' value='" & strTypeMaint & "'>" & vbCrLf

    Call Security_Selection_List("")
    If bolSecAutoSel Then
       Response.Write "<input type='submit' value='Select' onClick='goUrl(this.form.Url);return false;'>" & vbCrLf
    Else
       Response.Write "<input type='submit' value='Select'>" & vbCrLf
    End If
    Response.Write "</form>" & vbCrLf
    Response.Write "</td></tr>" & vbCrLf

    Call Security_Maintenance_Links(2)
    call Wrapup_Web_Page

%>

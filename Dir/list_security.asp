<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/security_select.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->
<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/security_maint_links.inc" -->

<%
    DIM intSub, strHld, bolExpFnd

    strSpecSecGrp = "SecMaint"
    Call System_Setup("NONE")
    Call Logon_Check(GetCurPath("")&"list_security.asp",5,strLogonGrp)

    bolBldSecurAryOnly = true

    strMaint = "Y"

    intCS = 7
    call Setup_Web_Page("Security Profiles",intCS)

    Call Security_Selection_List("")

    Response.Write "<tr><td><b>Name</b></td></td><td><b>User<br>Name</b></td><td><b>Email Address</b></td><td><b>Last<br>Access</b></td><td><b>Expiration</b></td><td><b>Security<br>Level</b></td><td align='center'><b>Security Groups</b></td></tr>" & vbCrLf
    Response.Write "<tr><td colspan='" & intCS & "'><hr align='left' size='5' noshade></td></tr>" & vbCrLf

    For intSub = 1 to arySecurity(0,0)
       Response.Write "<tr>" & vbCrLf

       Response.Write "<td>" & arySecurity(intSub,2) & "</td>" & vbCrLf
       Response.Write "<td>" & arySecurity(intSub,3) & "</td>" & vbCrLf
       Response.Write "<td>" & arySecurity(intSub,4) & "</td>" & vbCrLf
       strHld = arySecurity(intSub,8)
       If datediff("d", strHld, Date) > intActivityCheck Then
          strHld = "<font color='red'>" & strHld & "</font>"
          bolExpFnd = true
       End If
       Response.Write "<td>" & strHld & "</td>" & vbCrLf
       strHld = arySecurity(intSub,9)
       If datediff("d", Date, strHld) < 1 Then
          strHld = "<font color='red'>" & strHld & "</font>"
          bolExpFnd = true
       End If
       Response.Write "<td>" & strHld & "</td>" & vbCrLf
       Response.Write "<td align='center'>" & arySecurity(intSub,6) & "</td>" & vbCrLf
       Response.Write "<td>" & arySecurity(intSub,7) & "</td>" & vbCrLf
       Response.Write "<tr>" & vbCrLf
    Next

    Response.Write "<tr><td><p>&nbsp;</p></td></tr>" & vbCrLf

    call Wrapup_Web_Page

%>

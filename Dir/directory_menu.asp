<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->
<!--#INCLUDE FILE="../dir/directory_setup.inc" -->
<!--#INCLUDE FILE="../dir/form_data.inc" -->

<%
    DIM bolIncNotes, bolIncCompany, strOpts, bolSecAccess, strHld, strSubLevl, strSrch, strGroup, strDirOptionFn
    DIM strTitle, intRQL, strNewGrp

    strSrch = GetParmInfo("$S")
    strGroup = GetParmInfo("$G")
    strDirOptionFn = GetParmInfo("OFN")
    If lcase(strGroup) = lcase(Application("PGMDIR")) Then
       SESSION("HOMEDIR") = strGroup
       strGroup = ""
    End If
    If strGroup <> SESSION("GROUP") and strGroup <> SESSION("HOMEDIR") Then
       strNewGrp = ""
       If strGroup <> "" Then
          If CheckFileExist(Server.MapPath("/" & strGroup & "/index.htm")) <> 0 then
             ' Not from a Base Folder
             strNewGrp = strGroup
          Else
             SESSION("HOMEDIR") = strGroup
          End If
       End If
       If strNewGrp <> SESSION("GROUP") Then
          SESSION("GROUP") = trim(strNewGrp)
          Session.Contents.Remove("LOGONSECGRP")
       End If
    End If

    strSubLevl = lcase(RemoveSpaces(SESSION("SECGRP")))

    If SESSION("HOMEDIR") = "" Then
       SESSION("HOMEDIR") = mid(Request.ServerVariables("SCRIPT_NAME"),2)
       SESSION("HOMEDIR") = left(SESSION("HOMEDIR"),instr(SESSION("HOMEDIR"),"/")-1)
    End If

    If (strSubLevl <> "" and strSubLevl <> SESSION("SUBLEVL")) or SESSION("DBNAME") = "" or SESSION("NEWLOGON") = "Y" Then
       If  SESSION("NEWLOGON") = "Y" Then
          Session.Contents.Remove("MENUFN")
          Session.Contents.Remove("SUBLEVL")
       End If
       If strSubLevl <> SESSION("SUBLEVL") and strSubLevl <> "all" and strSubLevl <> "" Then
          SESSION("SUBLEVL") = strSubLevl
          strHld = RemoveSpaces(SESSION("SECGRP"))
          If CheckFileExist(Server.MapPath("/" & SESSION("HOMEDIR") & "/" & strHld & "_menu.asp")) = 0 then
             SESSION("SUBLEVL") = strHld
          End If
       End If
    End If

    Call Database_Setup

    If SESSION("NEWLOGON") = "Y" or SESSION("KWO") = ""  Then
       If strDirOptionFn = "" Then
          strDirOptionFn = GetParmInfo("$SG")
          If strDirOptionFn = "" Then
             strDirOptionFn = SESSION("SECGRP")
          End If
       End If
       Call Process_External_Keywords("@" & strDirOptionFn)
       SESSION("LOGONSECGRP") = BuildLogonSecGrp(strLogonGrp)
    End If

    If bolSecurityOnly Then
       intRQL = 5
    Else
       intRQL = 0
    End If

    Call System_Setup("NONE")
    Call Logon_Check(GetCurPath("")&"Directory_Menu.asp",intRQL,strLogonGrp)

    Session.Contents.Remove("$S")
    Session.Contents.Remove("$G")
    Session.Contents.Remove("$SG")
    Session.Contents.Remove("OFN")
    Session.Contents.Remove("SETUP")

    strOpts = "P"
    If instr(SESSION("KWO"),"C") > 0 Then
       bolIncCompany = true
       strOpts = strOpts & "C"
    End If

    If instr(SESSION("KWO"),"N") > 0 Then
       bolIncNotes = true
    End If

    If instr(SESSION("KWO"),"n") > 0 Then
       strOpts = strOpts & "N"
    End If

    If strSrch = "Y" and not bolSecurityOnly Then
       Response.Redirect("directory_list.asp?SPO=Y&DT=" & strOpts & "&E2C=Y&SC=Y")
    End If

    bolDisplayMsg = false

    strSpecSecGrp = "SecMaint"
    Call Verify_Group_Access("N/R","",5)
    If not bolSecError Then
       bolSecAccess = true
    End If

    If SESSION("GROUP") <> "" Then
       strTitle = SESSION("GROUP")
       if right(ucase(strTitle),1) <> "S" Then
          strTitle = strTitle & "'s"
       End If
    End If

%>

<html>
<head>
<%If bolSecurityOnly Then%>
<title>Security Maintenance Menu</title>
<%Else%>
<title><%=strTitle%> Directory Menu</title>
<%End If%>
</head>
<%
If Application("BGCLR") <> "" Then
   strHld = "<body bgcolor='" & Application("BGCLR") & "' link='Navy' vlink='Navy' alink='Navy' text='Black'"
Else
   strHld = "<body link='Navy' vlink='Navy' alink='Navy' text='Black' background='" & Application("BGIMG") & "'"
End If
Response.Write strHld & ">" & vbCrLf
%>
<table border="0" cellspacing="0" cellpadding="0" align="center">
<tr><td align='center'>
<font size='+2' color='#800000'>
<%If bolSecurityOnly Then%>
<b>Security Maintenance Menu</b>
<%Else%>
<b><%=strTitle%> Directory Menu</b>
<%End If%>
</font>
</td></tr>
<tr><td align="center">
<img src="/<%=Application("PGMDIR")%>/graphics/red_line.gif" height="3" border="0" alt="" width="400">
</td></tr>
<tr><td>
<p>&nbsp;</p>
</td></tr>
<%If Not bolSecurityOnly Then%>
<%If SESSION("SECLEVL") > 0 Then %>
<tr><td>
<img src="/<%=Application("PGMDIR")%>/graphics/spacer.gif" width="30" height="1" border="0" alt="">
<img src="/<%=Application("PGMDIR")%>/graphics/Red_Dot.gif" width="12" height="12" border="0" alt="">
&nbsp;&nbsp;
<font color="#800000"><b>List Options</b></font>
</td></tr>
<tr><td>
<img src="/<%=Application("PGMDIR")%>/graphics/spacer.gif" width="55" height="1" border="0" alt="">
<img src="/<%=Application("PGMDIR")%>/graphics/Red_Dot.gif" width="8" height="8" border="0" alt="">
&nbsp;&nbsp;
<a href='directory_list.asp?DT=P&NA=Y&ST=SG&E2C=Y&KW=@<%=SESSION("SUBLEVL")%>'><b>Name and Phone #</b></a>
</td></tr>
<tr><td>
<img src="/<%=Application("PGMDIR")%>/graphics/spacer.gif" width="55" height="1" border="0" alt="">
<img src="/<%=Application("PGMDIR")%>/graphics/Red_Dot.gif" width="8" height="8" border="0" alt="">
&nbsp;&nbsp;
<a href='directory_list.asp?DT=P&ST=SG&E2C=Y&KW=@<%=SESSION("SUBLEVL")%>'><b>Name, eMail and Phone #</b></a>
</td></tr>
<%If bolIncCompany Then%>
<tr><td>
<img src="/<%=Application("PGMDIR")%>/graphics/spacer.gif" width="55" height="1" border="0" alt="">
<img src="/<%=Application("PGMDIR")%>/graphics/Red_Dot.gif" width="8" height="8" border="0" alt="">
&nbsp;&nbsp;
<a href='directory_list.asp?DT=PC&ST=SG&E2C=Y&KW=@<%=SESSION("SUBLEVL")%>'><b>Name, eMail, Phone # and Company</b></a>
</td></tr>
<%End If%>
<%If bolIncNotes Then%>
<tr><td>
<img src="/<%=Application("PGMDIR")%>/graphics/spacer.gif" width="55" height="1" border="0" alt="">
<img src="/<%=Application("PGMDIR")%>/graphics/Red_Dot.gif" width="8" height="8" border="0" alt="">
&nbsp;&nbsp;
<a href='directory_list.asp?DT=PN&ST=SG&E2C=Y&KW=@<%=SESSION("SUBLEVL")%>'><b>Name, eMail, Phone # and Notes</b></a>
</td></tr>
<tr><td>
<img src="/<%=Application("PGMDIR")%>/graphics/spacer.gif" width="55" height="1" border="0" alt="">
<img src="/<%=Application("PGMDIR")%>/graphics/Red_Dot.gif" width="8" height="8" border="0" alt="">
&nbsp;&nbsp;
<a href='directory_list.asp?DT=N&ST=SG&E2C=Y&KW=@<%=SESSION("SUBLEVL")%>'><b>Name, eMail and Notes</b></a>
</td></tr>
<%End If%>
<%If bolIncNotes and bolIncCompany Then%>
<tr><td>
<img src="/<%=Application("PGMDIR")%>/graphics/spacer.gif" width="55" height="1" border="0" alt="">
<img src="/<%=Application("PGMDIR")%>/graphics/Red_Dot.gif" width="8" height="8" border="0" alt="">
&nbsp;&nbsp;
<a href='directory_list.asp?DT=A&ST=SG&E2C=Y&KW=@<%=SESSION("SUBLEVL")%>'><b>All Details</b></a>
</td></tr>
<%End If%>
<%If not bolSecAccess Then%>
<tr><td>&nbsp;</td></tr>
<%End If%>
<tr><td>
<img src="/<%=Application("PGMDIR")%>/graphics/spacer.gif" width="30" height="1" border="0" alt="">
<img src="/<%=Application("PGMDIR")%>/graphics/Red_Dot.gif" width="12" height="12" border="0" alt="">
&nbsp;&nbsp;
<font color="#800000"><b>Build Options</b></font>
</td></tr>
<tr><td>
<img src="/<%=Application("PGMDIR")%>/graphics/spacer.gif" width="55" height="1" border="0" alt="">
<img src="/<%=Application("PGMDIR")%>/graphics/Red_Dot.gif" width="8" height="8" border="0" alt="">
&nbsp;&nbsp;
<a href='directory_list.asp?DEF=E&DT=PC&ST=SG&E2C=Y&KW=@<%=SESSION("SUBLEVL")%>'><b>Build Email Mailing List</b></a>
</td></tr>
<tr><td>
<img src="/<%=Application("PGMDIR")%>/graphics/spacer.gif" width="55" height="1" border="0" alt="">
<img src="/<%=Application("PGMDIR")%>/graphics/Red_Dot.gif" width="8" height="8" border="0" alt="">
&nbsp;&nbsp;
<a href='directory_list.asp?DEF=W&DT=PC&ST=SG&E2C=Y&KW=@<%=SESSION("SUBLEVL")%>'><b>Build Web Page Display</b></a>
</td></tr>
<%End If%>
<%
strSpecSecGrp = "VCF"
Call Verify_Group_Access("N/R","",9)
If not bolSecError Then
%>
<tr><td>
<img src="/<%=Application("PGMDIR")%>/graphics/spacer.gif" width="55" height="1" border="0" alt="">
<img src="/<%=Application("PGMDIR")%>/graphics/Red_Dot.gif" width="8" height="8" border="0" alt="">
&nbsp;&nbsp;
<a href='directory_list.asp?DEF=V&DT=PC&ST=SG&E2C=Y&KW=@<%=SESSION("SUBLEVL")%>'><b>Build Web VCF List</b></a>
</td></tr>
<%End If%>
<%If not bolSecAccess Then%>
<tr><td>&nbsp;</td></tr>
<%End If%>
<%If SESSION("SECLEVL") > 2 Then %>
<tr><td>
<img src="/<%=Application("PGMDIR")%>/graphics/spacer.gif" width="30" height="1" border="0" alt="">
<img src="/<%=Application("PGMDIR")%>/graphics/Red_Dot.gif" width="12" height="12" border="0" alt="">
&nbsp;&nbsp;
<font color="#800000"><b>Membership Maintenance Options</b></font>
</td></tr>
<tr><td>
<img src="/<%=Application("PGMDIR")%>/graphics/spacer.gif" width="55" height="1" border="0" alt="">
<img src="/<%=Application("PGMDIR")%>/graphics/Red_Dot.gif" width="8" height="8" border="0" alt="">
&nbsp;&nbsp;
<a href='directory_maint.asp?TM=A'><b>Add Address</b></a>
</td></tr>
<tr><td>
<img src="/<%=Application("PGMDIR")%>/graphics/spacer.gif" width="55" height="1" border="0" alt="">
<img src="/<%=Application("PGMDIR")%>/graphics/Red_Dot.gif" width="8" height="8" border="0" alt="">
&nbsp;&nbsp;
<a href='directory_select.asp?MT=Y&TM=U'><b>Update Addresses</b></a>
</td></tr>
<tr><td>
<img src="/<%=Application("PGMDIR")%>/graphics/spacer.gif" width="55" height="1" border="0" alt="">
<img src="/<%=Application("PGMDIR")%>/graphics/Red_Dot.gif" width="8" height="8" border="0" alt="">
<%Else%>
<tr><td>
<img src="/<%=Application("PGMDIR")%>/graphics/spacer.gif" width="30" height="1" border="0" alt="">
<img src="/<%=Application("PGMDIR")%>/graphics/Red_Dot.gif" width="12" height="12" border="0" alt="">
<%End If%>
&nbsp;&nbsp;
<a href='<%=strDirHelpFn%>'><b>Documentation</b></a>
</td></tr>
<%End If%>
<%
If bolSecAccess Then
%>
<%If not bolSecAccess Then%>
<tr><td>&nbsp;</td></tr>
<%End If%>
<tr><td>
<img src="/<%=Application("PGMDIR")%>/graphics/spacer.gif" width="30" height="1" border="0" alt="">
<img src="/<%=Application("PGMDIR")%>/graphics/Red_Dot.gif" width="12" height="12" border="0" alt="">
&nbsp;&nbsp;
<font color="#800000"><b>Security Maintenance Options</b></font>
</td></tr>
<tr><td>
<img src="/<%=Application("PGMDIR")%>/graphics/spacer.gif" width="55" height="1" border="0" alt="">
<img src="/<%=Application("PGMDIR")%>/graphics/Red_Dot.gif" width="8" height="8" border="0" alt="">
&nbsp;&nbsp;
<a href='security_maint.asp?TM=A'><b>Add Security</b></a>
</td></tr>
<tr><td>
<img src="/<%=Application("PGMDIR")%>/graphics/spacer.gif" width="55" height="1" border="0" alt="">
<img src="/<%=Application("PGMDIR")%>/graphics/Red_Dot.gif" width="8" height="8" border="0" alt="">
&nbsp;&nbsp;
<a href='select_security.asp'><b>Update Security</b></a>
</td></tr>
<tr><td>
<img src="/<%=Application("PGMDIR")%>/graphics/spacer.gif" width="55" height="1" border="0" alt="">
<img src="/<%=Application("PGMDIR")%>/graphics/Red_Dot.gif" width="8" height="8" border="0" alt="">
&nbsp;&nbsp;
<a href='list_security.asp'><b>List Security Profiles</b></a>
</td></tr>
<tr><td>
<img src="/<%=Application("PGMDIR")%>/graphics/spacer.gif" width="55" height="1" border="0" alt="">
<img src="/<%=Application("PGMDIR")%>/graphics/Red_Dot.gif" width="8" height="8" border="0" alt="">
&nbsp;&nbsp;
<a href='Select_Security_Groups.asp'><b>Email Security Groups</b></a>
</td></tr>
<tr><td>
<img src="/<%=Application("PGMDIR")%>/graphics/spacer.gif" width="55" height="1" border="0" alt="">
<img src="/<%=Application("PGMDIR")%>/graphics/Red_Dot.gif" width="8" height="8" border="0" alt="">
&nbsp;&nbsp;
<a href='maint_text_file.asp?FN=[CD]_Addr.log '><b>View Access Log</b></a>
</td></tr>
<%End If%>
</table>
</body>
</html>

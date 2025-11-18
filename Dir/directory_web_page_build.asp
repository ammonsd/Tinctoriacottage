<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/directory_setup.inc" -->

<%
   Call System_Setup("NONE")
   Call Logon_Check(GetCurPath("")&"directory_web_page_build.asp",SESSION("RQL"),strLogonGrp)
%>

<html>
<head>
</head>
<body link='Blue' vlink='Blue' alink='Blue' onLoad="window.open('build_email_list.asp?NP=Y')">
<div align='center'>
<br><br><br><br>
<a href="<%=SESSION("MENUFN")%>"><b>Directory Menu</b></a>
<br><br>
<font size='-1' color='red'>
If the generated web page does not appear, your browser may be blocking popup windows.<br>
<%IF strBrowser = "Netscape" Then%>
To allow popups for this feature, select <b>Tools => Popup Manager => Allow Popups From This Site.</b>
<%ElseIf strBrowser = "IE" Then%>
To allow popups for this feature, press the &ltCtrl&gt key <b>while clicking</b> on the "Create Web Page" button.
<%Else%>
Review the <a href="<%=strDirHelpFn%>">documentation</a> for a workaround solution.
<%End If%>
</font>
</div>
</body></html>

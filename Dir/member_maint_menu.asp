

<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/system_setup.inc" -->
<!--#INCLUDE FILE="../dir/process_setup.inc" -->

<%

   DIM strBodyHtmlTag

   Call System_Setup("NONE")

   If Application("BGCLR") <> "" Then
      strBodyHtmlTag = "<body bgcolor='" & Application("BGCLR") & "' link='Navy' vlink='Navy' alink='Navy' text='Black'>"
   ElseIf Application("BGIMG") <> "" Then
      strBodyHtmlTag = "<body link='Navy' vlink='Navy' alink='Navy' text='Black' background='" & Application("BGIMG") & "'>"
   Else
      strBodyHtmlTag = "<body link='Navy' vlink='Navy' alink='Navy' text='Black'>"
   End If

%>
<html>
<head>

<title>Web Site Maintenance Menu</title>

</head>
<%=strBodyHtmlTag%>
<table border="0" cellspacing="0" cellpadding="0" align="center">
<tr><td align='center'>
<font size='+2' color='#800000'>
<b>Web Site Maintenance Menu</b>
</font>
</td></tr>
<tr><td align="center">
<img src="/dir/graphics/red_line.gif" height="3" border="0" alt="" width="400">
</td></tr>

<tr><td>&nbsp;</td></tr>
<tr><td>
<img src="/dir/graphics/spacer.gif" width="30" height="1" border="0" alt="">
<img src="/dir/graphics/Red_Dot.gif" width="12" height="12" border="0" alt="">
&nbsp;&nbsp;
<font color="#800000"><b>Membership Maintenance Options</b></font>
</td></tr>
<tr><td>
<img src="/dir/graphics/spacer.gif" width="55" height="1" border="0" alt="">
<img src="/dir/graphics/Red_Dot.gif" width="8" height="8" border="0" alt="">
&nbsp;&nbsp;
<a href='/dir/member_maint.asp?TM=A'><b>Add Member</b></a>
</td></tr>
<tr><td>
<img src="/dir/graphics/spacer.gif" width="55" height="1" border="0" alt="">
<img src="/dir/graphics/Red_Dot.gif" width="8" height="8" border="0" alt="">
&nbsp;&nbsp;
<a href='/dir/member_select.asp'><b>Update Member</b></a>
</td></tr>
<tr><td>
<img src="/dir/graphics/spacer.gif" width="55" height="1" border="0" alt="">
<img src="/dir/graphics/Red_Dot.gif" width="8" height="8" border="0" alt="">
&nbsp;&nbsp;
<a href='/dir/members_list.asp?SRCH=Y'><b>List Members</b></a>
</td></tr>
<tr><td>
<img src="/dir/graphics/spacer.gif" width="55" height="1" border="0" alt="">
<img src="/dir/graphics/Red_Dot.gif" width="8" height="8" border="0" alt="">
&nbsp;&nbsp;
<a href='/dir/members_broadcast_email.asp'><b>Broadcast Email Members</b></a>
</td></tr>

</table>
</body>
</html>


<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/box_it.inc" -->
<!--#INCLUDE FILE="../dir/system_setup.inc" -->
<!--#INCLUDE FILE="../dir/process_setup.inc" -->

<%
CALL System_Setup("NONE")
CALL Box_Top_Section
Session.Contents.Remove("PSWD")
SESSION("P") = SESSION("MENUFN")
%>

<H3><B>You are now logged off.</B></H3><BR>
The information entered during this session remains in the Web browser's memory until cleared from the browser's cache or the browser is closed.
To protect the confidentiality of the information make sure the browser's disk cache is cleared to prevent someone
else from being able to view the information temporarily stored on the computer.
This can be accomplished by either
<ul>
<li>
Shutting the browser down which will, in effect, clear the cache, or
</li>
<li>
Clear the cache using the instructions provided in the browser's online help system.
</li>
</ul>
<br><br>
<div align="center">
<a href="<%=strLogonLoc%>logon.asp><img src="/<%=Application("PGMDIR")%>/graphics/login_sm.gif" alt="Logon" width="46" height="17" border="0">
</a>
</div>
</td>

<%
CALL Box_Bottom_Section
%>

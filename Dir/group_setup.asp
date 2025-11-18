<%@ LANGUAGE="VBSCRIPT" %>
<%
   If SESSION("SETUP") <> "Y" Then

      SESSION("HOMEDIR") = mid(Request.ServerVariables("SCRIPT_NAME"),2)
      SESSION("HOMEDIR") = left(SESSION("HOMEDIR"),instr(SESSION("HOMEDIR"),"/")-1)

       'User Friendly Name of Security Group
      If Request.QueryString("G") <> "" Then
         SESSION("$G") = Request.QueryString("G")
      End If

       'Name of the "DIR" options file
      If Request.QueryString("OFN") <> "" Then
         SESSION("OFN") = Request.QueryString("OFN")
      End If
      If SESSION("OFN") = "" Then
         SESSION("OFN") = SESSION("$G")
      End If

       'Security Access Group
      If Request.QueryString("SG") <> "" and ucase(Request.QueryString("SG")) <> ucase(Application("PGMDIR")) Then
         SESSION("$SG") = Request.QueryString("SG")
         SESSION("HOMEDIR") = SESSION("$SG")
      End If

      If Request.QueryString("S") <> "" Then
          ' Y - Display Search Window
          ' N - Display Menu (default)
         SESSION("$S") = Request.QueryString("S")
      End If

      SESSION("SETUP") = "Y"
      Response.Redirect "/" & Application("PGMDIR") & "/Directory_Menu.asp"

   Else
      Response.Redirect "/" & Application("PGMDIR") & "/Directory_Menu.asp"
   End If
%>

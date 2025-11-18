
<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
Server.ScriptTimeout = 99999
%>

<!--#INCLUDE FILE="../dir/maint_webpage.inc" -->
<!--#INCLUDE FILE="../dir/process_setup.inc" -->
<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/get_data.inc" -->
<!--#INCLUDE FILE="../dir/select_entry.inc" -->

<%

   DIM objJMail, objUpload, File, Item, strMsg, strServer, strDir, intCntr, strEMail, strUplFns

   Call System_Setup("NONE")
   strSpecSecGrp = "FileUpload"
   Call Logon_Check("upload.asp",5,strLogonGrp)

   strGetData = SESSION("GETDATA")
   Session.Contents.Remove("GETDATA")

   strServer = Request.ServerVariables("SERVER_NAME")
   strServer = mid(strServer,instr(strServer,".")+1)

   strEMail = Request.QueryString("EM")
   strDir = Request.QueryString("DIR")

   If Request.QueryString("UPL") <> "Y" Then
      If strEMail = "" Then
         strEMail = Application("SUPPORT_MAIL")
         If strEMail = "" Then
            strEMail = "admin@" & strServer
         End If
      End If
      If strDir = "" Then
         If  Request.Form("DIR") <> "" Then
            strDir = Request.Form("DIR")
         ElseIf strGetData <> "Y" Then
            If SESSION("EMPGM") <> "" Then
               strDir = GetSelection(Application("EMUPLDIR"),"DIR","Select Upload Directory","upload.asp")
            End If
            If strDir = "" Then
               strDir = GetSelection(Application("UPLDIR"),"DIR","Select Upload Directory","upload.asp")
            End If
         End If
      End If
   End If

   If strGetData = "Y" and strDir = "" Then
      intCntr = 0
      If SESSION("EMPGM") <> "" Then
         Response.Redirect(SESSION("EMPGM"))
      Else
         Call Ending_Message
      End If
      Response.END
   End If

   If strDir = "" Then
      strDir = "/cgi-bin/"
   End If
   If left(strDir,1) <> "/" Then
      strDir = "/" & strDir
   End If
   If right(strDir,1) <> "/" Then
      strDir = strDir & "/"
   End If

   If Request.QueryString("UPL") <> "Y" Then
      Call Display_Form(strDir,strEMail)
      Response.End
   End If

   Set objUpload = Server.CreateObject("Persits.Upload.1")
   ' Upload files
   objUpload.SetMaxSize 20971520 ' Truncate files above 20MB
   objUpload.SaveVirtual strDir

   If SESSION("EMPGM") <> "" Then
      If instr(SESSION("EMPGM"),"?") > 0 Then
         For Each File in objUpload.Files
             If strUplFns = "" Then
                strUplFns = strDir & File.ExtractFileName
             Else
                strUplFns = strUplFns & "," & strDir & File.ExtractFileName
             End If
         Next
         If strUplFns = "" Then
            strUplFns = "?" & strDir
         End If
         SESSION("EMPGM") = SESSION("EMPGM") & strUplFns
      End If
      Set objUpload = Nothing
      Response.Redirect(SESSION("EMPGM"))
      Response.END
   End If

   Set objJMail = Server.CreateObject( "JMail.SMTPMail" )
   objJMail.ServerAddress = "mail." & strServer & ":25"
   objJMail.Sender = SESSION("USEREMAIL")
   objJMail.Subject = "Files Uploaded to " & strServer & strDir
   objJMail.AddRecipient(strEMail)
   strMsg = vbCrLf
   intCntr = 0
   For Each File in objUpload.Files
       intCntr = intCntr + 1
       strMsg = strMsg & intCntr & ") " & File.ExtractFileName & vbCrLf
   Next
   strMsg = strMsg & vbCrLf & vbCrLf & "Uploaded By: " & SESSION("USERNAME") & " (" & SESSION("USEREMAIL") & ")"
   objJMail.Body = strMsg
   If intCntr > 0 Then
      objJMail.Execute()
   End If
   Set objJMail = Nothing
   Set objUpload = Nothing

   Response.Write "<br><br><div align='center'><b>" & vbCrLf
   If intCntr = 0 Then
      Response.Write "No Files Were Uploaded" & vbCrLf
   Else
      Response.Write "Requested Files Have Been Successfully Uploaded to the Web Site and the Owner Has Been Notified" & vbCrLf
   End If
   Response.Write "</b></div>" & vbCrLf

SUB Ending_Message

   Response.Write "<br><br><div align='center'><b>" & vbCrLf
   If intCntr = 0 Then
      Response.Write "No Files Were Uploaded" & vbCrLf
   Else
      Response.Write "Requested Files Have Been Successfully Uploaded to the Web Site and the Owner Has Been Notified" & vbCrLf
   End If
   Response.Write "</b></div>" & vbCrLf
END SUB

SUB Display_Form(DIR,EM)

   DIM strBodyHtmlTag

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
<title></title>
</head>
<%=strBodyHtmlTag%>
<table align='center' border='0' cellpadding='' cellspacing='' bgcolor='#b0c4de'>
<tr><td align='center' colspan='2'>
<font size='+1' color='#990000'>
<b>Upload Files</b>
</font>
</td></tr>
<tr><td align='center'>
<img src='/dir/graphics/red_line.gif' width=150 height='3' border='0' alt=''>
</td></tr>

<tr><td>
<p>&nbsp;</p>
</td></tr>
<form method="post" enctype="multipart/form-data" action="upload.asp?UPL=Y&DIR=<%=DIR%>&EM=<%=EM%>">
<tr><td>
<input type='file' name='File1' size='80' maxlength='80'>
</td></tr>
<tr><td>
<input type='file' name='File2' size='80' maxlength='80'>
</td></tr>
<tr><td>
<input type='file' name='File3' size='80' maxlength='80'>
</td></tr>
<tr><td>
<input type='file' name='File4' size='80' maxlength='80'>
</td></tr>
<tr><td>
<input type='file' name='File5' size='80' maxlength='80'>
</td></tr>
<tr><td>
<input type='file' name='File6' size='80' maxlength='80'>
</td></tr>
<tr><td>
<input type='file' name='File7' size='80' maxlength='80'>
<tr><td>
<input type='file' name='File8' size='80' maxlength='80'>
</td></tr>
<tr><td>
<input type='file' name='File9' size='80' maxlength='80'>
</td></tr>
<tr><td>
<input type='file' name='File10' size='80' maxlength='80'>
</td></tr>
<tr><td colspan='2'>
<hr width='100%' size='5' noshade>
</td></tr>
<tr><td align='center'>
<input type='SUBMIT' value='Upload'>
</td></tr>
</table>
</body></html>
<%END SUB%>


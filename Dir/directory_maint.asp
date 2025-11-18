<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/end_msg.inc" -->
<!--#INCLUDE FILE="../dir/maint_webpage.inc" -->
<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/directory_select.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->
<!--#INCLUDE FILE="../dir/validate_email_format.inc" -->
<!--#INCLUDE FILE="../dir/display_text_file.inc" -->
<!--#INCLUDE FILE="../dir/form_data.inc" -->
<!--#INCLUDE FILE="../dir/cookie_maint.inc" -->

<%
   DIM strDbAction, intErrNbr, strTemp, strInterActive, bolDbInactive
   DIM strRecID, strMaintType, strName, strHld, strAddedEmail, X

   strTypeGoto = "NF"
   bolDbInactive = false 'Not Used

   If Request.QueryString("BPG") Then
      SESSION("BPG") = Request.QueryString("BPG")
   End if

   strMaintType = FormData("TM")
   If strMaintType = "" Then
      strMaintType = "A"
   End If

   strRecID = FormData("ID")

   strMaint = "Y"
   Call System_Setup("directory_maint.asp")
   Call Logon_Check(GetCurPath("")&"directory_maint.asp",3,strLogonGrp)

   Call Database_Setup

   Set errMsgs = CreateObject("Scripting.Dictionary")

   strDbAction = left(ucase(FormData("AC")),1)
   strInterActive = FormData("IA")

   strFName = trim(ReplSpecChars(FormData("First Name")))
   strLName = trim(ReplSpecChars(FormData("Last Name")))
   strEmail = trim(lcase(ReplSpecChars(FormData("Email"))))
   strKeyWords = trim(ReplSpecChars(FormData("Keywords")))
   strContactInfo = trim(ReplSpecChars(FormData("Contact Info")))
   strCompany = trim(ReplSpecChars(FormData("Company")))
   strSecGrps = trim(FormData("Security Groups"))

   If Request.Form("PROCESS") <> "Y" Then
      If strDbAction = "A" and strInterActive = "Y" Then
         call Display_Form
         Response.END
      ElseIf strDbAction="" Then
         strDbAction = strMaintType
         If strDbAction = "U" Then
            Call Get_DB_Record(strRecID)
         End If
         call Display_Form
         Response.END
      End If
   ElseIf strDbAction <> "D" Then
      Call Validate_Entries
   End If

   If errMsgs.Count > 0 Then
      aryErrorMsgs = errMsgs.Items
      err.clear
      call Process_Errors
   Else
      If strDbAction = "D" then
         Call DeleteRecord(Request.Form("ID"))
      ElseIf strDbAction = "U" then
         Call UpdateRecord(Request.Form("ID"))
      ElseIf strDbAction = "C" then
         strHld = "directory_maint.asp?AC=A&IA=Y"
         strHld = strHld & "&Last%20Name=" & ChgSpecChars(Request.Form("Last Name"))
         strHld = strHld & "&First%20Name=" & ChgSpecChars(Request.Form("First Name"))
         strHld = strHld & "&email=" & lcase(ChgSpecChars(trim(Request.Form("Email"))))
         strHld = strHld & "&Keywords=" & ChgSpecChars(Request.Form("Keywords"))
         strHld = strHld & "&Contact%20Info=" & ChgSpecChars(Request.Form("Contact Info"))
         strHld = strHld & "&Company=" & ChgSpecChars(Request.Form("Company"))
         strHld = strHld & "&Security%20Groups=" & ChgSpecChars(Request.Form("Security Groups"))
         strHld = ReplParmChars(strHld)
         Response.Redirect(strHld)
      Else
         Call Add_Record
      End If
      bolIncLinks = true
      If errMsgs.Count > 0 Then
         aryErrorMsgs = errMsgs.Items
         err.clear
         call Process_Errors
      ElseIf SESSION("BPG") <> "" Then
         If strDbAction <> "D" then
            Call AddCookie("DIRMAINT", "SECGROUPS", strSecGrps, 365)
         End If
         CALL GoBack_Pages
      Else
      End If
   End If

   call Display_Form

FUNCTION UpdateRecord(ID)

   DIM intSub

   Call Setup_Global_Record_Details("U")

   aryDbDtls(0) = ID

   If UpdateGlobalRecord Then
      intNbrPgs = 2
      strBackName = "Return to Selection"
      SESSION("POPUPMSG") = "Y"
      call Display_Ending_Msg("Details Successfully Updated","")
   End If

END FUNCTION

SUB Add_Record

   DIM intSub

   Call Setup_Global_Record_Details("A")

   If AddGlobalRecord Then
      SESSION("POPUPMSG") = "Y"
      call Display_Ending_Msg("New Entry Successfully Added" ,"")
      strName = trim(strFName & " " & strLName)
   End If

END SUB

FUNCTION DeleteRecord(ID)

   Call Setup_Global_Record_Details("D")

   aryDbDtls(0) = ID

   If DeleteGlobalRecord Then
      intNbrPgs = 2
      strBackName = "Return to Selection"
      call Display_Ending_Msg("Entry Successfully Deleted","")
   End If

END FUNCTION

FUNCTION ReplSpecChars(DATA)
   ReplSpecChars = ReplChars(DATA,"^",CHR(13))
   ReplSpecChars = ReplChars(ReplSpecChars,CHR(149),"#")
   ReplSpecChars = ReplChars(ReplSpecChars,CHR(183),"&")
   ReplSpecChars = ReplChars(ReplSpecChars,CHR(186),"?")
END FUNCTION

FUNCTION ChgSpecChars(DATA)
   ChgSpecChars = ReplChars(DATA,CHR(13),"^")
   ChgSpecChars = ReplChars(ChgSpecChars,"#",CHR(149))
   ChgSpecChars = ReplChars(ChgSpecChars,"&",CHR(183))
   ChgSpecChars = ReplChars(ChgSpecChars,"?",CHR(186))
END FUNCTION

SUB Display_Form

   DIM intCntr, strHld

   strTblClr = "#b0c4de"
   strFSize = "+1"

   If strDbAction = "A" then
      Call Setup_Web_Page("Add New Entry",2)
   Else
     If strInterActive = "Y" Then
        Call Setup_Web_Page("Review New Entry Details",2)
     Else
        Call Setup_Web_Page("Update Details",2)
     End If
   End If

%>
<form action="directory_maint.asp" method="post" name="FORM" id="FORM">
<%If strDbAction = "A" and strInterActive = "Y" Then
Else%>
<input type="hidden" name="OrigEmail" value=<%=streMail %>>
<%End If%>
<input type="hidden" name="ID" value=<%=strRecID %>>
<input type="hidden" name="PROCESS" value="Y">
<tr><td>
<b>First Name:</b>
</td><td>
<input type="text" name="First Name" size="50" maxlength="50" value="<%=strFName %>">
</td></tr>
<tr><td>
<b>Last Name:</b>
</td><td>
<input type="text" name="Last Name" size="50" maxlength="50" value="<%=strLName %>">
</td></tr>
<tr><td>
<b>eMail Address:</b>
</td><td>
<input type="text" name="Email" size="50" maxlength="50" value="<%=strEmail %>">
</td></tr>
<%
If strDbAction = "A" Then
   If not bolEditError Then
      If not bolSecError and ucase(SESSION("SECGRP")) = "ALL" then
         strSecGrps = GetCookie("DIRMAINT","SECGROUPS")
      Else
         strSecGrps = ""
      End If
      If strSecGrps = "" and ucase(SESSION("SECGRP")) <> "ALL" then
         strSecGrps = SESSION("SECGRP")
      End If
   End If
End If
strSpecSecGrp = "SecMaint"
bolDisplayMsg = false
Call Verify_Group_Access("N/R","",5)
If bolSecError Then%>
<input type="hidden" name="Security Groups" value="<%=strSecGrps %>">
<%Else%>
<tr><td>
<b>Security Groups:</b>
&nbsp;&nbsp;
</td><td nowrap>
<input type="text" name="Security Groups" size="50" maxlength="50" value="<%=strSecGrps %>">
</td></tr>
<%End If%>
<tr><td>
Company:
</td><td>
<input type="text" name="Company" size="50" maxlength="50" value="<%=strCompany %>">
</td></tr>
<tr><td>
Contact Info:
</td><td nowrap>
<textarea cols='38' rows='6' name='Contact Info'><%=strContactInfo%></textarea>
</td></tr>
<tr><td>
Keywords:
</td><td nowrap>
<textarea cols='38' rows='6' name='Keywords'><%=strKeyWords%></textarea>
</td></tr>
<tr><td width="100%" colspan="2">
<hr width="100%" size="5" noshade>
</td></tr>
<tr><td colspan="2" align="center">
<%If strDbAction = "A" then %>
<input type="submit" name="AC" value="Add">
<%Else%>
<input type="submit" name="AC" value="Update">
&nbsp;&nbsp;&nbsp;&nbsp;
<input type="submit" name="AC" value="Copy">
<%End If
If not bolSecError Then%>
&nbsp;&nbsp;&nbsp;&nbsp;
<input type="button"  value="Available Groups" onClick="window.open('directory_security_groups.asp')">
<%End If%>
<%If strDbAction <> "A" then
strSpecSecGrp = "DirDelete"
Call Verify_Group_Access("N/R","",5)
If not bolSecError or SESSION("USERID") = lcase(strAddedUserID) Then%>
&nbsp;&nbsp;&nbsp;&nbsp;
<input type="submit" name="AC" value="Delete" onCLick="return confirm('You are about to Delete the Current Entry -- Do You Wish to Continue?')">
<%End If
  End If%>
&nbsp;&nbsp;&nbsp;&nbsp;
<%If strDbAction = "A" then %>
<input type='reset' value='Clear Entry'>
<%Else%>
<input type='reset' value='Reset'>
<%End If%>
</form>
</td></tr>
<%
call Wrapup_Web_Page
End Sub
%>

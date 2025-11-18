<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/check_date.inc" -->
<!--#INCLUDE FILE="../dir/write_errors.inc" -->
<!--#INCLUDE FILE="../dir/end_msg.inc" -->
<!--#INCLUDE FILE="../dir/security_maint_links.inc" -->
<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/security_select.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->
<!--#INCLUDE FILE="../dir/popup_email.inc" -->
<!--#INCLUDE FILE="../dir/validate_email_format.inc" -->
<!--#INCLUDE FILE="../dir/cookie_maint.inc" -->
<!--#INCLUDE FILE="../dir/display_text_file.inc" -->

<%

   DIM strLookUpKey, bolRecFnd, strDbAction, intErrNbr, strTemp, strRecID, strMaintType
   DIM strPswd2, strHld

   strMaintType = FormData("TM")
   If strMaintType = "" Then
      strMaintType = "A"
   End If

   strRecID = FormData("LID")

   strMaint = "Y"
   bolLogAccess = true
   bolPwUpdAllowed = true

   strSpecSecGrp = "SecMaint"
   Call System_Setup("NONE")
   Call Logon_Check(GetCurPath("")&"security_maint.asp",5,strLogonGrp)

   Call Database_Setup
   CALL Get_Function_Security_Groups

   strDbAction = ucase(FormData("AC"))

   If strDbAction="" Then
      strDbAction = strMaintType
      If strDbAction = "U" Then
         Call Get_Security(strRecID)
      End If
      call Display_Form
      Response.END
   End If

   strDbAction = left(strDbAction,1)

   Set errMsgs = CreateObject("Scripting.Dictionary")

   If strDbAction <> "D" Then
      Call Validate_Input
   End If

   If errMsgs.Count > 0 Then
      aryErrorMsgs = errMsgs.Items
      err.clear
      call Process_Errors
   Else
      If strDbAction = "D" then
         bolRecFnd = DeleteRecord(Request.Form("LID"))
      ElseIf strDbAction = "U" then
         bolRecFnd = UpdateRecord(Request.Form("LID"))
      Else
         Call Add_Security_Record
      End If
      bolIncLinks = true
      strTypeLink = "S"
      If errMsgs.Count > 0 Then
         aryErrorMsgs = errMsgs.Items
         err.clear
         call Process_Errors
      ElseIf Err.Number <> 0 and Err.Number <> 53 Then
        call Write_Errors
    ElseIf strDbAction = "D" then
         intNbrPgs = 2
         strBackName = "Return to Selection"
         call Display_Ending_Msg("Profile Successfully Deleted","")
      End If
   End If

   call Display_Form

SUB Validate_Input

  '// Validate form entries.

   strSecFName = Request.Form("First Name")
   strSecLName = Request.Form("Last Name")
   strUserName = lcase(Request.Form("User Name"))
   strSecEMail = lcase(Request.Form("Email"))
   intSecLevel = Request.Form("Security Level")
   intUserExpDate = Request.Form("User Exp Date")
   strSecSecGrps = trim(Request.Form("Security Groups"))

   Call Validate_Security_Entry

End SUB

FUNCTION UpdateRecord(ID)

   Set objRecordSet = Server.CreateObject("ADODB.Recordset")
   objRecordSet.Open APPLICATION("SECURTBL"), strDbConn, adOpenStatic, adLockPessimistic, adCmdTable

   If objRecordSet.EOF Then
       UpdateRecord = false
   Else
     objRecordSet.MoveFirst
     objRecordSet.Find ("ID = " & ID)
     If objRecordSet.EOF Then
        UpdateRecord = false
     Else
        UpdateRecord = true
        objRecordSet.Fields("Updated") = NOW()
        If datediff("d", objRecordSet.Fields("Last Access"), Date) > intActivityCheck Then
           objRecordSet.Fields("Last Access") = NOW()
        End If
        objRecordSet.Fields("Update User ID") = SESSION("USERID")
        objRecordSet.Fields("First Name") = strSecFName
        objRecordSet.Fields("Last Name") = strSecLName
        objRecordSet.Fields("User Name") = strUserName
        objRecordSet.Fields("Email") = strSecEMail
        If Request.Form("ResetPSWD") = "Y" Then
           strPswd = "Pass+word"
           objRecordSet.Fields("Password") = Encrypt(strPswd,strUserName)
        End If
        If Request.Form("ExpPswd") = "on" or Request.Form("ResetPSWD") = "Y" Then
           objRecordSet.Fields("Password Exp Date") = dateadd("d", -1, Date)
           bolExpPswd = true
        Else
           objRecordSet.Fields("Password Exp Date") = dateadd("d", int(intPswdExpCycle), Date)
        End If
        objRecordSet.Fields("Security Level") = intSecLevel

        objRecordSet.Fields("Security Groups") = FormatSecurityGroups(strSecSecGrps)

        If Request.Form("LogAccess") = "Y" Then
           bolLogAccess = true
        Else
           bolLogAccess = false
        End If
        objRecordSet.Fields("LogAccess") = bolLogAccess
        objRecordSet.Fields("Logon Attempts") = 0

        If Request.Form("Allow Pswd Update") = "Y" Then
           bolPwUpdAllowed = true
        Else
           bolPwUpdAllowed = false
        End If
        objRecordSet.Fields("Allow Pswd Update") = bolPwUpdAllowed

        objRecordSet.Fields("Password Exp Cycle") = intPswdExpCycle

        If intUserExpDate = "" Then
           objRecordSet.Fields("User Exp Date") = "12/31/9999"
        ElseIf instr(intUserExpDate,"/") = 0 Then
           objRecordSet.Fields("User Exp Date") = dateadd("d", int(intUserExpDate), Date)
        Else
           objRecordSet.Fields("User Exp Date") = intUserExpDate
        End If

        bolIncLinks = true
        strTypeLink = "S"
        intNbrPgs = 2
        strBackName = "Return to Selection"
        SESSION("POPUPMSG") = "Y"
        call Display_Ending_Msg("Security Profile Successfully Updated","")
     End If
   End If

   objRecordSet.Update
   objRecordSet.Close
   Set objRecordSet = Nothing

END FUNCTION

FUNCTION DeleteRecord(LogonID)

   Set objRecordSet = Server.CreateObject("ADODB.Recordset")
   objRecordSet.ActiveConnection = strDbConn
   objRecordSet.Source = "DELETE FROM " & APPLICATION("SECURTBL") & " WHERE [ID] = " & LogonID
   objRecordSet.LockType = adLockReadOnly
   objRecordSet.Open
   Set objRecordSet = Nothing

END FUNCTION

SUB Security_Notification(TYP,DIR)

   DIM strSubj, intLoc

   IF SESSION("SECLEVL") < 5 Then
      intNbrPgs = 2
      strBackName = "Return to Selection"
      SESSION("POPUPMSG") = "Y"
      call Display_Ending_Msg("Security Profile Successfully Updated","")
      EXIT SUB
   End If

   aryText(1,1) = trim(strSecFName)
   If aryText(1,1) = "" Then
      aryText(1,1) = trim(strSecLName)
   End If
   If aryText(1,1) = "" Then
      aryText(1,1) = strUserName
   End If
   aryText(2,1) = strUserName
   If strSecEMail = "" Then
      aryText(3,1) = "<font color='red'><b>NONE</b></font>"
   Else
      aryText(3,1) = strSecEMail
   End If

   aryText(4,1) = strPswd

   aryText(5,1) = SupportDistribution()
   aryText(6,1) = GetPgmPath("")
   aryText(7,1) = trim(DIR)

   If intSecLevel > 2 Then
      aryText(10,1) = "[SKIP]"
   Else
      aryText(9,1) = "[SKIP]"
   End If

   strMsgVar = "MSG"
   Call Display_Text_File("security_details.txt")
   If strDisSubj <> "" Then
      strSubj = strDisSubj
      intLoc = instr(strSubj,"[SB]")
      If intLoc > 0 Then
         strSubj = left(strSubj,intLoc-1) & DIR & mid(strSubj,intLoc + 4)
      End If
   End If
   If strSecEMail <> "" Then
      strToAddr = strSecEMail
      SESSION("EMN") = strSecEMail & " " & ReplParmChars(strSecFName & "%20" & strSecLName)
      SESSION("SUBJ") = strSubj
      Call Popup_eMail(strToAddr,"","",strSubj,"** Overlay This Line with Text Copied from Browser **")
   End If

END SUB

FUNCTION FormData(FldName)

   If Request.QueryString(FldName) <> "" Then
      FormData = Request.QueryString(FldName)
   Else
      FormData = Request.Form(FldName)
   End If

END FUNCTION

SUB Display_Form

   DIM intCntr, strHld

   strTblClr = "#b0c4de"
   strFSize = "+1"

   If strDbAction = "A" then
      Call Setup_Web_Page("Create New Account",2)
      If not bolEditError Then
         strSecSecGrps = GetCookie("SECMAINT","SECGROUPS")
         intSecLevel = GetCookie("SECMAINT","SECLEVEL")
         If GetCookie("SECMAINT","LOGACCESS") = "Y" Then
            bolLogAccess = true
         End If
         If GetCookie("SECMAINT","PSWDUPDOK") = "Y" Then
            bolPwUpdAllowed = true
         End If
         intPswdExpCycle = GetCookie("SECMAINT","EXPPSWD")
         intUserExpDate = GetCookie("SECMAINT","EXPDATE")
      End If
   Else
      Call Setup_Web_Page("Update Account Details",2)
   End If

   If int(intPswdExpCycle) < 1 Then
      intPswdExpCycle = 30
   End If

%>
<form action="security_maint.asp" method="post" name="FORM" id="FORM">
<input type="hidden" name="FM" value="<%=strUserName %>">
<input type="hidden" name="OrigUName" value=<%=strUserName %>>
<input type="hidden" name="LID" value=<%=strRecID %>>
<%If strOrigPswd <> "" Then%>
<input type="hidden" name="OrigPswd" value=<%=strOrigPswd %>>
<%End If%>
<tr><td>
<b>First Name:</b>
</td><td>
<input type="text" name="First Name" size="50" maxlength="50" value="<%=strSecFName %>">
</td></tr>
<tr><td>
<b>Last Name:</b>
</td><td>
<input type="text" name="Last Name" size="50" maxlength="50" value="<%=strSecLName %>">
</td></tr>
<tr><td>
<b>User Name:</b>
</td><td>
<input type="text" name="User Name" size="50" maxlength="50" value="<%=strUserName %>">
</td></tr>
<tr><td>
<b>Email Address:</b>
</td><td>
<input type="text" name="Email" size="50" maxlength="50" value="<%=strSecEMail %>">
</td></tr>
<tr><td>
<b>Security Level:</b>
</td><td nowrap>
<%Call Security_Level_Selection(intSecLevel)%>
</td></tr>
<tr><td>
<b>Password Cycle:</b>
</td><td nowrap>
<input type="text" name="Password Exp Cycle" size="4" maxlength="4" value="<%=intPswdExpCycle %>">
&nbsp;&nbsp;&nbsp;
Log Access:
<%IF bolLogAccess Then
  strHld = "checked"
Else
  strHld = ""
End If%>
<input type='checkbox' name='LogAccess' value='Y' <%=strHld%>>
</td></tr>
<tr><td>
<b>User Expiration:</b>
</td><td nowrap>
<input type="text" name="User Exp Date" size="15" maxlength="10" value="<%=intUserExpDate %>">
&nbsp;&nbsp;&nbsp;
Allow Password Changes:
<%IF bolPwUpdAllowed Then
  strHld = "checked"
Else
  strHld = ""
End If%>
<input type='checkbox' name='Allow Pswd Update' value='Y' <%=strHld%>>
</td></tr>
<%IF strDbAction <> "A" Then%>
<tr><td>
Reset Password:
</td><td nowrap>
<input type='checkbox' name='ResetPSWD' value='Y'>
&nbsp;&nbsp;&nbsp;
Expire Password:
<%IF bolExpPswd Then
  strHld = "checked"
Else
  strHld = ""
End If%>
<input type='checkbox' name='ExpPswd' value='on' <%=strHld%>>
</td></tr>
<%End If%>
<tr><td>
Security Groups:
</td><td nowrap>
<input type="text" name="Security Groups" size="50" maxlength="50" value="<%=strSecSecGrps %>">
</td></tr>
<tr><td width="100%" colspan="2">
<hr width="100%" size="5" noshade>
</td></tr>
<tr><td colspan="2" align="center">
<%If strDbAction = "A" then %>
<input type="submit" name="AC" value="Add Profile">
<%Else%>
<input type="submit" name="AC" value="Update Profile">
&nbsp;&nbsp;&nbsp;&nbsp;
<input type="submit" name="AC" value="Delete Profile" onCLick="return confirm('You are about to Delete a Security Profile -- Do You Wish to Continue?')">
<%End If%>
<%If not bolSecurityOnly and APPLICATION("DIRTBL") <> "" Then%>
&nbsp;&nbsp;&nbsp;&nbsp;
<input type="button"  value="Assigned Groups" onClick="window.open('select_security_groups.asp?L=Y')">
<%End If%>
&nbsp;&nbsp;&nbsp;&nbsp;
<%If strDbAction = "A" then %>
<input type='reset' value='Clear Entry'>
<%Else%>
<input type='reset' value='Reset Entry'>
<%End If%>
</form>
</td></tr>
<%Call Security_Maintenance_Links(2)%>
</table>
</div>

</body>
</html>

<%
End Sub
%>

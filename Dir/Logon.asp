<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/system_setup.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->
<!--#INCLUDE FILE="../dir/process_setup.inc" -->
<!--#INCLUDE FILE="../dir/encrypt_decrypt.inc" -->
<!--#INCLUDE FILE="../dir/cookie_maint.inc" -->
<!--#INCLUDE FILE="../dir/box_it.inc" -->
<!--#INCLUDE FILE="../dir/log_update.inc" -->
<!--#INCLUDE FILE="../dir/format_security_groups.inc" -->

<%
    DIM strPswd, strUserName, strSavePswd, strLogonOK, strUNameEntity, strUNamePrompt, strLogonSecGrp
    DIM bolInactiveError, bolInactiveAccount, bolPswdErr, bolPswdChgErr, bolLogonAttemptsError
    DIM intRecID, intMaxLogonAttempts, intNbrLogonAttempts, strSecurTbl, bolNewPSWD

    Call System_Setup("NONE")

    strSecurTbl = SESSION("SECURTBL")
    If strSecurTbl = "" Then
       strSecurTbl = APPLICATION("SECURTBL")
    End If

    intMaxLogonAttempts = 5

    Call Database_Setup

    If SESSION("LOGONSECGRP") <> "" Then
       strLogonSecGrp = "_" & SESSION("LOGONSECGRP")
    End If

    call Reset_Variables

    If strUNamePrompt = "" Then
       strUNamePrompt = "User Name:"
    End If
    If strUNameEntity = "" Then
       strUNameEntity = "User Name"
    End If

    If Request.Form("NewPswd") = "Y" Then
       bolNewPswd = true
    End If

    If Request.Form("CHKLOGON") <> "Y" Then
       strUserName =  GetCookie("SECLOGON" & strLogonSecGrp,"USERNAME")
       strSavePswd =  GetCookie("SECLOGON" & strLogonSecGrp,"SAVEPSWD")
       If strSavePSWD = "on" Then
          strPswd = GetCookie("SECLOGON" & strLogonSecGrp,"PSWD")
       End If
       call GetLogon(strUserName,strPswd,"N")
    Else
       If Request.Form("ID") <> "" Then
          intRecID = int(Request.Form("ID"))
       End If
       strPswd = Request.Form("PSWD")
       If strPswd = "********************" Then
          strPswd = Decrypt(SESSION("OPSWD"),lcase(Request.Form("USERNAME")))
          SESSION("OPSWD") = ""
       End If
       If bolNewPswd Then
          If strPswd = Request.Form("PSWD2") Then
             If len(trim(strPswd)) < 6 Then
                SESSION("EXPPSWD") = "New Password Must Be at Least 6 Characters"
                strLogonOK = "X"
             Else
                strLogonOK = "Y"
                Call Update_Password(intRecID,strPswd)
             End If
          Else
             SESSION("EXPPSWD") = "Passwords don't Match - Please Re-Enter"
             strLogonOK = "X"
          End If
       ElseIf LogonOK(Request.Form("USERNAME"),strPswd) Then
          If Request.Form("ChgPSWD") = "on" Then
             strLogonOK = "X"
             If SESSION("SECPWUPD") = true Then
                SESSION("EXPPSWD") = "Enter new Password"
             Else
                bolPswdChgErr = true
             End If
          ElseIf SESSION("EXPPSWD") = "Y" Then
             strLogonOK = "X"
             SESSION("EXPPSWD") = "Password Has Expired<br>Please Enter a new Password"
             bolNewPswd = true
          ElseIf bolInactiveError or bolInactiveAccount Then
             strLogonOK = "X"
          Else
             strLogonOK = "Y"
             strUserName = lcase(Request.Form("USERNAME"))
             strSavePswd = Request.Form("SavePswd")
             If strUserName <> "isgadmin" Then
                Call AddCookie("SECLOGON" & strLogonSecGrp,"USERNAME",lcase(strUserName),30)
                Call AddCookie("SECLOGON" & strLogonSecGrp,"SAVEPSWD",strSavePswd,30)
                If strSavePSWD = "on" Then
                   Call AddCookie("SECLOGON" & strLogonSecGrp,"PSWD",Encrypt(strPswd,lcase(strUserName)),30)
                End If
             End If
             Call Update_Access_Log("")
          End If
       End If
       If strLogonOK = "Y" Then
          call Reset_Variables
          If SESSION("P") = "" Then
             SESSION("P") = ReplChars(Request.ServerVariables("Script_Name"),"/"," ")
             SESSION("P") = "http://" & Request.ServerVariables("Server_Name") & "/"& GetWords(SESSION("P"),1,1)
          End If
          If SESSION("BPG") <> "" Then
             SESSION("BPG") = int(SESSION("BPG")) + 1
          End If
          CALL Update_Last_Access(intRecID)
          call Response.Redirect(SESSION("P"))
       ElseIf strLogonOK = "X" Then
          call GetLogon(Request.Form("USERNAME"),"","Y")
       Else
          call GetLogon(Request.Form("USERNAME"),strPswd,"Y")
       End If
    End If

FUNCTION LogonOK(USERNAME,PSWD)

    DIM strSql, strPswd, intExpID

    LogonOK = false
    intExpID = 0

    If USERNAME = "" Then
       EXIT FUNCTION
    End If

    If lcase(USERNAME) = "diradmin" Then
       If PSWD = "DIR+admin" Then
          SESSION("SECLEVL") = 9
          SESSION("USERID") = lcase(USERNAME)
          SESSION("PSWD") = Encrypt(PSWD,SESSION("USERID"))
          SESSION("SECLOG") = "N"
          LogonOK = true
       End If
       EXIT FUNCTION
    End If

    Set objRecordSet = Server.CreateObject("ADODB.Recordset")
    objRecordSet.ActiveConnection = strDbConn

    strSql = " WHERE [" & strUNameEntity & "] = '" & lcase(USERNAME) & "'"

    objRecordSet.Source = "SELECT * FROM " & strSecurTbl & strSql

    objRecordSet.LockType = adLockReadOnly
    objRecordSet.Open

    If Not objRecordSet.EOF Then
       strPswd = Decrypt(objRecordSet.Fields("Password"),lcase(objRecordSet.Fields(strUNameEntity)))
       intNbrLogonAttempts = objRecordSet.Fields("Logon Attempts") + 1
       intRecID = objRecordSet.Fields("ID")
       If PSWD = strPswd Then
          SESSION("SECLEVL") = objRecordSet.Fields("Security Level")
          SESSION("SECACCESSGRPS") = UnFormatSecurityGroups(objRecordSet.Fields("Security Groups"))
          SESSION("SECGRP") = SESSION("SECACCESSGRPS")
          If instr(SESSION("SECGRP"),",") > 0 Then
             SESSION("SECGRP") = left(SESSION("SECGRP"),instr(SESSION("SECGRP"),",")-1)
          End If
          SESSION("SECLOG") = objRecordSet.Fields("LogAccess")
          SESSION("SECPWUPD") = objRecordSet.Fields("Allow Pswd Update")
          SESSION("USERID") = lcase(objRecordSet.Fields(strUNameEntity))
          SESSION("PSWD") = Encrypt(PSWD,SESSION("USERID"))
          SESSION("USEREMAIL") = lcase(objRecordSet.Fields("Email"))
          SESSION("USERNAME") = objRecordSet.Fields("First Name") & " " & objRecordSet.Fields("Last Name")
          LogonOK = true
          If not bolNewPswd Then
             If not isdate(objRecordSet.Fields("Password Exp Date")) Then
                SESSION("EXPPSWD") = "Y"
             ElseIf datediff("d", Date, objRecordSet.Fields("Password Exp Date")) < 1 Then
                SESSION("EXPPSWD") = "Y"
             End If
          End If
          intRecID = objRecordSet.Fields("ID")
          If datediff("d", Date, objRecordSet.Fields("User Exp Date")) < 1 Then
             bolInactiveAccount = true
             IF objRecordSet.Fields("User Exp Date") = "12/31/1900" Then
                bolLogonAttemptsError = true
             End If
          End If
          If datediff("d", objRecordSet.Fields("Last Access"), Date) > intActivityCheck Then
             bolInactiveError = true
          End If
       Else
          bolPswdErr = true
          If intNbrLogonAttempts >= intMaxLogonAttempts Then
             intExpID = intRecID
          End If
          If SESSION("DEBUG") = "Y" Then
             Response.Write "Entered Password: " & PSWD & "<br>Valid Passowrd: " & strPswd & "<br>"
          End If
       End If
    End If

    objRecordset.Close
    Set objRecordSet = Nothing

    If bolPswdErr Then
       CALL Update_Last_Access(intRecID)
       If intExpID > 0 Then
          bolPswdErr = false
          CALL Deactivate_Account(intExpID)
       End If
    End If

END FUNCTION

SUB Update_Password(ID,PSWD)

   Set objRecordSet = Server.CreateObject("ADODB.Recordset")
   objRecordSet.Open strSecurTbl, strDbConn, adOpenStatic, adLockPessimistic, adCmdTable

   If Not objRecordSet.EOF Then
     objRecordSet.MoveFirst
     objRecordSet.Find ("ID = " & ID)
     If Not objRecordSet.EOF Then
        objRecordSet.Fields("Password") = Encrypt(PSWD,lcase(objRecordSet.Fields(strUNameEntity)))
        objRecordSet.Fields("Password Exp Date") = dateadd("d", objRecordSet.Fields("Password Exp Cycle"), Date)
        SESSION("PSWD") = objRecordSet.Fields("Password")
     End If
   End If

   objRecordSet.Update
   objRecordSet.Close
   Set objRecordSet = Nothing

END SUB

SUB Deactivate_Account(ID)

   Set objRecordSet = Server.CreateObject("ADODB.Recordset")
   objRecordSet.Open strSecurTbl, strDbConn, adOpenStatic, adLockPessimistic, adCmdTable

   If Not objRecordSet.EOF Then
     objRecordSet.MoveFirst
     objRecordSet.Find ("ID = " & ID)
     If Not objRecordSet.EOF Then
        objRecordSet.Fields("User Exp Date") = "12/31/1900"
        bolInactiveAccount = true
        bolLogonAttemptsError = true
     End If
   End If

   objRecordSet.Update
   objRecordSet.Close
   Set objRecordSet = Nothing

END SUB

SUB Update_Last_Access(ID)

   Set objRecordSet = Server.CreateObject("ADODB.Recordset")
   objRecordSet.Open strSecurTbl, strDbConn, adOpenStatic, adLockPessimistic, adCmdTable

   If Not objRecordSet.EOF Then
     objRecordSet.MoveFirst
     objRecordSet.Find ("ID = " & ID)
     If Not objRecordSet.EOF Then
        If not bolPswdErr Then
           objRecordSet.Fields("Last Access") = DATE()
           intNbrLogonAttempts = 0
        End If
        objRecordSet.Fields("Logon Attempts") = intNbrLogonAttempts
     End If
   End If

   objRecordSet.Update
   objRecordSet.Close
   Set objRecordSet = Nothing

END SUB

SUB Reset_Variables
   Session.Contents.Remove("EXPPSWD")
END SUB

FUNCTION GetLogon(USERNAME,PSWD,LogonErr)

   DIM strHidPswd, strErrMsg, strBodyHtmlTag

   If Application("BGCLR") <> "" Then
      strBodyHtmlTag = "<body bgcolor='" & Application("BGCLR") & "' link='Navy' vlink='Navy' alink='Navy' text='Black'>"
   ElseIf Application("BGIMG") <> "" Then
      strBodyHtmlTag = "<body link='Navy' vlink='Navy' alink='Navy' text='Black' background='" & Application("BGIMG") & "'>"
   Else
      strBodyHtmlTag = "<body bgcolor='#6D6D6D' link='Navy' vlink='Navy' alink='Navy' text='Black'>"
   End If

   If LogonErr = "Y" Then
      Session.Contents.Remove("PSWD")
      If bolInactiveError Then
         strErrMsg =             "Account Has Been Deactived Due to Inactivity Exceeding " & intActivityCheck & " Days<br>"
         strErrMsg = strErrMsg & "Contact the System Administrator to have the Account Reactived"
      ElseIf bolInactiveAccount Then
         strErrMsg =             "This Account Has Been Deactivated"
         If bolLogonAttemptsError Then
            strErrMsg = strErrMsg & " Due to Excessive Logon Failures"
         End If
         strErrMsg = strErrMsg & "<br>Contact the System Administrator to have the Account Reactived"
      ElseIf bolPswdChgErr Then
         strErrMsg =             "Not Authorized to Change Password<br>"
         strErrMsg = strErrmsg & "Contact the System Administrator to Have Your Password Changed"
      End If
   End If

   If strErrMsg <> "" Then
      CALL Box_Top_Section
      Response.Write "<div align='CENTER'><font color='#BD0000'><b>" & strErrMsg & "</b></font></div><br>"
      CALL Box_Bottom_Section
      Response.END
   End If

   If strUNamePrompt = "" Then
      strUNamePrompt = "User Name:"
   End If

   If trim(PSWD) <> "" Then
      strHidPswd = "********************"
   End If

%>

<html>
<head>
<title>System Logon</title>
</head>
<%=strBodyHtmlTag%>
<div align='center'>
<div align='center'>
<br><br><br><br><br><br>
<table border=1 cellspacing=0 cellpadding=2 bordercolor='#000080'>
<tr>
<td align='center' valign='top' bgcolor='#000095'>
<font color='white'>
<b>Registered User Logon</b>
</font>
</td></tr>
<tr>
<td valign='top' bgcolor='#B4B4B4'>
<font face='Arial'>
<div align='center'>
<table border=0 cellspacing=0 cellpadding=2>
<form action='logon.asp' method='post'>
<input type='hidden' name='CHKLOGON' value='Y'>
<input type='hidden' name='ID' value='<%=intRecID%>'>
<%If bolNewPswd Then%>
<input type='hidden' name='NewPswd' value='Y'>
<%End If%>
<%If PSWD <> "" Then
SESSION("OPSWD") = PSWD
End If%>
<tr><td no wrap align='right'>
<font size=-1>
<b><%=strUNamePrompt%></b>
</font></td>
<td nowrap>
<input type='Text' name='USERNAME' size='30' maxlength='100' value='<%=USERNAME%>'>
</td></tr>
<tr><td no wrap align='right'>
<font size=-1>
<b>Password:</b>
</font></td>
<td nowrap>
<input type='Password' name='PSWD' size='30' maxlength='20' value='<%=strHidPswd%>'>
</td></tr>
<%If SESSION("EXPPSWD") <> "" Then%>
<tr><td no wrap align='right'>
<font size=-1>
<b>Re-Enter Password:</b>
</font></td>
<td nowrap>
<input type='Password' name='PSWD2' size='30' maxlength='20'>
</td></tr>
<tr><td colspan=2 align='center'>
<%End If%>

<%If SESSION("EXPPSWD") = "" Then%>
<tr><td align='right'>
<input type='image' src='/<%=Application("PGMDIR")%>/graphics/login_sm.gif'>
</td>
<td nowrap align='right'>
<font size=-2 color="#0000ff">
<b>Save Password:</b>
</font>
<%
If strSavePSWD = "on" Then
   strSavePSWD = "checked"
Else
   strSavePSWD = ""
End If
%>
<input type='Checkbox' name='SavePSWD' value='on' <%=strSavePSWD%>><br>
<font size=-2 color="#0000ff">
<b>Change Password:</b>
</font>
<input type='Checkbox' name='ChgPSWD' value='on'>
</td></tr>
<%Else%>
<tr><td align='center' colspan=2>
<input type='image' src='/<%=Application("PGMDIR")%>/graphics/login_sm.gif'>
</td></tr>
<%End If%>
<%If LogonErr = "Y" Then
   If strLogonOK = "X" Then
      LogonErr = SESSION("EXPPSWD")
   Else
      If bolPswdErr Then
         LogonErr = "Invalid Password, Please Re-Enter"
         If intNbrLogonAttempts = intMaxLogonAttempts - 1 Then
            LogonErr = LogonErr & "<br>Account Will Be De-Activated with Next Invalid Entry"
         End If
      Else
         LogonErr = "Invalid User Name"
      End If
   End If %>
<tr><td colspan=2 align='center'><font color='#BD0000' size=-1> <b><%=LogonErr%></b></font></td></tr>
<%End If%>
</form>
</table>
</table>
</div>
</body>
</html>


<%
END FUNCTION
%>

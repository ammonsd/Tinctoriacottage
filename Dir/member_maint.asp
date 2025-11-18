<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/maint_webpage.inc" -->
<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/member_select.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->
<!--#INCLUDE FILE="../dir/form_data.inc" -->
<!--#INCLUDE FILE="../dir/display_edit_errors.inc" -->
<!--#INCLUDE FILE="../dir/check_numeric.inc" -->
<!--#INCLUDE FILE="../dir/end_msg.inc" -->
<!--#INCLUDE FILE="../dir/send_email.inc" -->
<!--#INCLUDE FILE="../dir/security_select.inc" -->

<%

   DIM strLookUpKey, strDbAction, intErrNbr, strTemp, strRecID, strMaintType, strHld, strGB, bolDbInactive

   strMaintType = FormData("TM")
   If strMaintType = "" Then
      strMaintType = "A"
   End If

   strRecID = FormData("ID")

   Call System_Setup("NONE")

   strDbAction = ucase(FormData("AC"))
   If strDbAction = "CLEAR ENTRY" Then
      strDbAction = ""
      strMaintType = "A"
   End If
   strDbAction = left(strDbAction,1)

   strGB = FormData("GB")

   If strGB <> "" Then
      strMaintType = "A"
      strDbAction = "A"
      If SESSION("USERID") = "" Then
         SESSION("USERID") = "guestbook"
      End If
   Else
      Call Logon_Check(GetCurPath("")&"member_maint.asp",3,strLogonGrp)
   End If

   Call Database_Setup
   Set errMsgs = CreateObject("Scripting.Dictionary")

   If Request.Form("PROCESS") <> "Y" and strGB = "" Then
      If strDbAction="" Then
         strDbAction = strMaintType
         If strDbAction = "U" Then
            Call Get_DB_Record(strRecID)
         End If
         call Display_Form
         Response.END
      End If
   ElseIf strDbAction <> "D" Then
      Call Validate_Input
   End If

   If errMsgs.Count > 0 Then
      aryErrorMsgs = errMsgs.Items
      err.clear
      If strGB <> "" Then
         call Display_Errors
      Else
         call Process_Errors
      End If

   Else
      If strDbAction = "D" then
         Call DeleteRecord(Request.Form("ID"))
      ElseIf strDbAction = "U" then
         Call UpdateRecord(Request.Form("ID"))
      Else
         Call Add_Record
         If errMsgs.Count = 0 and strGB <> "" Then
            Call New_Member_Notification
            If APPLICATION("BOOKLOGON") = "Y" and trim(Request.Form("UserName")) <> "" Then
               Call Create_Member_Logon
            End If
         End If
      End If
      If errMsgs.Count > 0 Then
         aryErrorMsgs = errMsgs.Items
         err.clear
         If strGB <> "" Then
            call Display_Errors
         Else
            call Process_Errors
         End If
      End If
   End If

   If strGB <> "" Then
      Response.Redirect(strGB)
   Else
      call Display_Form
   End If

SUB Validate_Input

  '// Validate form entries.


   DIM strMailingList

   strFName = trim(Request.Form("FirstName"))
   strLName = trim(Request.Form("LastName"))
   strEmail = trim(Request.Form("e-Mail"))
   strStreet = trim(Request.Form("Street"))
   strCity = trim(Request.Form("City"))
   strState = trim(Request.Form("State"))
   strZip = trim(Request.Form("Zip"))
   strCountry = trim(Request.Form("Country"))
   strComments = trim(Request.Form("Comments"))
   strMailingList = trim(Request.Form("e-Mail List"))
   strKeywords = trim(Request.Form("Keywords"))

   If strMailingList = "on" Then
      bolMailingList = true
      bolDbInactive = false
   Else
      bolMailingList = false
      bolDbInactive = true
   End If

   Call Validate_Entries

End SUB

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
      If strGB <> "" Then
         strHld = "Thank You "
         If strFName <> "" Then
            strHld = strHld & strFName & " "
         End If
         strHld = strHld & "For Updating Our Guest Book"
         call Display_Ending_Msg(strHld,"")
      Else
         call Display_Ending_Msg("New Entry Successfully Added" ,"")
      End If
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

SUB New_Member_Notification

   DIM strEmailMsg, strAddress, strHld

   strAddress = ""
   If strStreet <> "" Then
      strAddress = strAddress & strStreet & CHR(10)
   End If
   strHld = " "
   If strCity <> "" Then
      strAddress = strAddress & strCity
      strHld = ", "
   End If
   If strState <> "" Then
      strAddress = strAddress & strHld & strState
   End If
   If strZip <> "" Then
      strAddress = strAddress & " " & strZip
   End If
   If strCountry <> "United States" Then
      If strAddress <> "" and strStreet <> "" Then
         strAddress = strAddress & CHR(10)
      End If
      strAddress = strAddress & strCountry
   End If

   strEmailMsg = trim(strFName & " " & strLName)
   If Request.Form("UserName") <> "" Then
      strEmailMsg = strEmailMsg & " (" & Request.Form("UserName") & ")"
   End If
   strEmailMsg = strEmailMsg & CHR(10)
   strEmailMsg = strEmailMsg & strEmail & CHR(10)
   strEmailMsg = strEmailMsg & strAddress
   If strComments <> "" Then
      strEmailMsg = strEmailMsg & CHR(10) & CHR(10) & strComments
   End If

   strEmailMsg = strEmailMsg & CHR(10) & CHR(10)
   If bolMailingList Then
      strEmailMsg = strEmailMsg & "I"
   Else
      strEmailMsg = strEmailMsg & "Do Not i"
   End If
   strEmailMsg = strEmailMsg & "nclude me on your Mailing List."
   call SendEmail(strSMTPServer,strSupportMail,trim(strFName & " " & strLName),strEmail, "New Guestbook Entry",strEmailMsg)

END SUB

SUB Create_Member_Logon

   strSecFName = strFName
   strSecLName = strLName
   strSecEMail = strEMail
   strUserName = Request.Form("UserName")
   strPSWD = Request.Form("Password")
   intSecLevel = 1
   strSecSecGrps = "BOOKSPG"
   intPswdExpCycle = 9999
   strUserID = "guestbook"

   If strFName <> "" and strLName <> "" and strEMail <> "" and strUserName <> "" and strPSWD <> "" Then
      Call Add_Security_Record
   End If

END SUB

SUB Display_Form

   DIM intCntr, strHld

   If strDbAction = "A" then
      strHld = "Create New Member Entry"
   Else
      strHld = "Update Member Details"
   End If


   bolCountryReq = true

   strTblClr = "#b0c4de"
   strFSize = "+1"
   Call Setup_Web_Page(strHld,2)

%>
<form action="member_maint.asp" method="post">
<input type="hidden" name="PROCESS" value="Y">
<%If strRecID <> "" Then%>
<input type="hidden" name="ID" value='<%=strRecID%>'>
<input type="hidden" name="OrigEmail" value=<%=streMail %>>
<%End If%>
<%If strGB <> "" Then%>
<input type="hidden" name="GB" value='<%=strGB%>'>
<%End If%>
<tr>
<td><b>First Name</b></td>
<td>
<input type="text" name="FirstName" size="60" maxlength="50" value='<%=strFName%>'>
</td>
</tr>
<tr><td><b>Last Name</b></td><td>
<input type="text" name="LastName" size="60" maxlength="50" value='<%=strLName%>'>
</td>
</tr>
<tr>
<td><b>Email address</b></td>
<td>
<input type="text" name="e-Mail" size="60" maxlength="50" value='<%=strEmail%>'>
</td>
</tr>
<tr>
<td>Street Address</td>
<td>
<input type="text" name="Street" size="60" maxlength="50" value='<%=strStreet%>'>
</td>
</tr>
<tr>
<td>City </font></td>
<td>
<input type="text" name="City" size="60" maxlength="50" value='<%=strCity%>'>
</td>
</td>
<tr>
<td>State</td>
<td>
<%call Build_States_Select_List(strState)%>
Zip
<input type="text" name="Zip" size="5" maxlength="10" value='<%=strZip%>'>
</td>
</tr>
<tr>
<td>Country</td>
<td>
<%call Build_Country_Select_List(strCountry)%>
</td>
</tr>
<tr>
<td>e-Mail List</td>
<td>
<%If bolMailingList or strDbAction = "A" then
   strHld = " checked"
Else
   strHld = ""
End If%>
<input type="Checkbox"<%=strHld%> name="e-Mail List">
</td>
</tr>
<tr><td>
Notes:
</td><td nowrap>
<textarea cols='45' rows='4' name='Keywords'><%=strKeyWords%></textarea>
</td></tr>
<tr><td width='100%' colspan='2'>
<hr width='100%' size='5' noshade>
</td></tr>
<tr><td colspan='2' align='center'>
<%If strDbAction = "A" then %>
<input type='submit' name='AC' value='Add Member'>
<%Else%>
<input type='submit' name='AC' value='Update Member'>
&nbsp;&nbsp;&nbsp;&nbsp;
<input type='submit' name='AC' value='Delete Member' onCLick="return confirm('You are about to Delete this Member -- Do You Wish to Continue?')">
<%End If%>
&nbsp;&nbsp;&nbsp;&nbsp;
<%If strDbAction = "A" then %>
<input type='submit' name='AC' value='Clear Entry'>
<%Else%>
<input type='reset' value='Reset Entry'>
<%End If%>
</form>
</td></tr>
<%Call Maintenance_Links(2)%>
</table>
</div>

</body>
</html>

<%
End Sub
%>

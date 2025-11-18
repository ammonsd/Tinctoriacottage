<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/maint_webpage.inc" -->
<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/book_select.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->
<!--#INCLUDE FILE="../dir/form_data.inc" -->
<!--#INCLUDE FILE="../dir/end_msg.inc" -->
<!--#INCLUDE FILE="../dir/cookie_maint.inc" -->

<%

   DIM strLookUpKey, strDbAction, intErrNbr, strTemp, strRecID, strMaintType, strHld

   strMaintType = FormData("TM")
   If strMaintType = "" Then
      strMaintType = "A"
   End If

   strRecID = FormData("ID")

   Call System_Setup("NONE")
   Call Logon_Check(GetCurPath("")&"book_maint.asp",3,strLogonGrp)

   Call Database_Setup

   strDbAction = left(ucase(FormData("AC")),1)

   Set errMsgs = CreateObject("Scripting.Dictionary")

   If Request.Form("PROCESS") <> "Y" Then
      If strDbAction="" Then
         strDbAction = strMaintType
         If strDbAction = "U" Then
            bolNoFormat = true
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
      call Process_Errors
   Else
      If strDbAction = "D" then
         Call DeleteRecord(Request.Form("ID"))
      ElseIf strDbAction = "U" then
         Call UpdateRecord(Request.Form("ID"))
      Else
         Call Add_Record
      End If
      If errMsgs.Count > 0 Then
         aryErrorMsgs = errMsgs.Items
         err.clear
         call Process_Errors
      ElseIf strDbAction <> "D" then
         Call AddCookie("BOOKMAINT", "CATEGORY", strCategory, 365)
         Call AddCookie("BOOKMAINT", "LANG", strLang, 365)
      End If
   End If

   call Display_Form

SUB Validate_Input

  '// Validate form entries.

   DIM intSub

   intErrNbr = 0

   strTitle = trim(Request.Form("Title"))
   strAuthor = trim(Request.Form("Author"))
   strDetails = trim(Request.Form("Details"))
   strCategory = trim(Request.Form("Category"))
   strPicFile = lcase(trim(Request.Form("Picture File")))
   strISBN = trim(Request.Form("ISBN"))
   strCrDate = trim(Request.Form("CrDate"))
   strLang = trim(Request.Form("Lang"))

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

FUNCTION DeleteRecord(ID)

   Call Setup_Global_Record_Details("D")

   aryDbDtls(0) = ID

   If DeleteGlobalRecord Then
      intNbrPgs = 2
      strBackName = "Return to Selection"
      call Display_Ending_Msg("Entry Successfully Deleted","")
   End If

END FUNCTION

SUB Add_Record

   Call Setup_Global_Record_Details("A")

   If AddGlobalRecord Then
      intNbrPgs = 1
      strBackName = "Return"
      SESSION("POPUPMSG") = "Y"
      call Display_Ending_Msg("New Entry Successfully Added" ,"")
   End If

END SUB

SUB Display_Form

   DIM intCntr, strHld

   If strDbAction = "A" then
      strHld = "Create New Book Entry"
      If not bolEditError Then
         strCategory = GetCookie("BOOKMAINT","CATEGORY")
         strLang = GetCookie("BOOKMAINT","LANG")
      End If
   Else
      strHld = "Update Book Details"
   End If

   strMaint = "Y"
   strTblClr = "#b0c4de"
   strFSize = "+1"
   Call Setup_Web_Page(strHld,2)

%>
<form action="book_maint.asp" method="post">
<input type='hidden' name='ID' value=<%=strRecID %>>
<input type="hidden" name="PROCESS" value="Y">
<tr><td>
<b>Title:</b>
</td><td>
<%If instr(strTitle,"'") = 0 Then
   strHld = "'" & strTitle & "'"
Else
  strHld = CHR(34) & strTitle & CHR(34)
End If%>
<input type='text' name='Title' size='95' maxlength='100' value=<%=strHld%>>
</td></tr>
<tr><td>
<b>Author:</b>
</td><td>
<%If instr(strAuthor,"'") = 0 Then
   strHld = "'" & strAuthor & "'"
Else
  strHld = CHR(34) & strAuthor & CHR(34)
End If%>
<input type='text' name='Author' size='95' maxlength='100' value=<%=strHld%>>
</td></tr>
<tr><td>
Picture File:
</td><td>
<input type='text' name='Picture File' size='95' maxlength='100' value='<%=strPicFile %>'>
</td></tr>
<tr><td>
<b>Category:</b>
</td><td>
<%call Build_Category_Selection(strCategory)%>
&nbsp;&nbsp;
Copyright:
&nbsp;&nbsp;&nbsp;
<input type='text' name='CrDate' size='4' maxlength='4' value='<%=strCrDate %>'>
&nbsp;&nbsp;
ISBN:
&nbsp;&nbsp;&nbsp;
<input type='text' name='ISBN' size='13' maxlength='20' value='<%=strISBN %>'>
&nbsp;&nbsp;
Language
<input type='text' name='LANG' size='15' maxlength='15' value='<%=strLang %>'>
</td></tr>
<tr><td>
<b>Status:</b>
</td><td>
<%IF not bolInactive Then
  strHld = "checked"
Else
  strHld = ""
End If%>
Active
<input type='Radio' name='Status' value='A' <%=strHld%>>
<%IF bolInactive Then
  strHld = "checked"
Else
  strHld = ""
End If%>
Inactive
<input type='Radio' name='Status' value='I' <%=strHld%>>
</td></tr>
<tr><td colspan=2 align='center'>
<b>Details</b>
</td></tr>
<tr><td colspan=2 align='center'>
<textarea name='Details' cols='84' rows=14 wrap='YES'>
<%=strDetails%>
</textarea>
</td></tr>
<tr><td width='100%' colspan='2'>
<hr width='100%' size='5' noshade>
</td></tr>
<tr><td colspan='2' align='center'>
<%If strDbAction = "A" then %>
<input type='submit' name='AC' value='Add Book'>
<%Else%>
<input type='submit' name='AC' value='Update Book'>
&nbsp;&nbsp;&nbsp;&nbsp;
<input type='submit' name='AC' value='Delete Book' onCLick="return confirm('You are about to Delete a Book Entry -- Do You Wish to Continue?')">
<%End If%>
&nbsp;&nbsp;&nbsp;&nbsp;
<%If strDbAction = "A" then %>
<input type='reset' value='Clear Entry'>
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

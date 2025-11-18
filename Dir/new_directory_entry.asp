
<% @ LANGUAGE="VBSCRIPT" %>
<%
 OPTION EXPLICIT
 Response.Buffer = true

%>

<!--#INCLUDE FILE="../dir/maint_webpage.inc" -->
<!--#INCLUDE FILE="../dir/directory_setup.inc" -->
<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/send_email.inc" -->
<!--#INCLUDE FILE="../dir/write_errors.inc" -->
<!--#INCLUDE FILE="../dir/get_data.inc" -->
<!--#INCLUDE FILE="../dir/directory_select.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->
<!--#INCLUDE FILE="../dir/check_dupe_email.inc" -->

<%

   DIM strToName, strToEmail, strFrEmail, strFrName, strMsg, strSubj, strEndMsg, bolRR, intSub, intErrNbr, strHld

   Call System_Setup("NONE")
   Call Logon_Check(GetCurPath("")&"new_directory_entry.asp",1,strLogonGrp)

   strGetData = SESSION("GETDATA")
   Session.Contents.Remove("GETDATA")

   If strGetData <> "Y" Then
      call Display_Form
   End If

   call Database_Setup

   Set errMsgs = CreateObject("Scripting.Dictionary")

   Call Validate_Input
   If errMsgs.Count > 0 Then
      call Process_Errors
      call Display_Form
      Response.End
   End If

   strToName = trim(Request.Form("SN"))
   strToEmail = trim(Request.Form("SE"))
   If strToName = "" Then
      strToName = strSupportName
   End If
   If strToEmail = "" Then
      strToEmail = strSupportMail
   End If

   strFrName = trim(Request.Form("FrN"))
   strFrEmail = trim(Request.Form("FrEM"))
   If strFrName = "" Then
      strFrName = SESSION("USERNAME")
   End If
   If strFrEmail = "" Then
      strFrEmail = SESSION("USEREMAIL")
   End If

   strSubj = trim(Request.Form("Subj"))

   strMsg = "<TEXT>"
   strMsg = strMsg & "1) Select and Copy the following text to C:\TEMP\" & Replace(trim(Request.Form("FName") & " " & Request.Form("LName"))," ","_") & ".vcf" & CHR(10)
   strMsg = strMsg & "2) Run VCF_2_Addr_Dir to create shortcut for adding new entry to directory" & CHR(10) & CHR(10) & CHR(10)
   strMsg = strMsg & "BEGIN:VCARD" & CHR(10)
   strMsg = strMsg & "VERSION:2.1" & CHR(10)
   strMsg = strMsg & "N:" & Request.Form("LName") & ";" & Request.Form("FName") & CHR(10)
   strMsg = strMsg & "FN:" & Request.Form("LName") & "," & Request.Form("FName") & CHR(10)
   If Request.Form("ORG") <> "" Then
      strMsg = strMsg & "ORG:" & Request.Form("ORG") & CHR(10)
   End If
   If Request.Form("PHONE") <> "" Then
      strMsg = strMsg & "TEL;WORK:" & Request.Form("PHONE") & CHR(10)
   End If
   If Request.Form("EMAIL") <> "" Then
      strMsg = strMsg & "EMAIL:" & Request.Form("EMAIL") & CHR(10)
   End If
   If Request.Form("JTitle") <> "" Then
      strMsg = strMsg & "TITLE:" & Request.Form("JTitle") & CHR(10)
   End If
   If Request.Form("NOTE") <> "" Then
      strMsg = strMsg & "NOTE:" & Request.Form("NOTE") & CHR(10)
   End If

   strMsg = strMsg & "END:VCARD" & CHR(10)

   strEndMsg = "Email Request Failed"
   If strFrEmail <> "" and strFrName<> "" and strToName <> "" and strToEmail <> "" and strSubj <> "" Then
      strToEmail = strToName & " <" & strToEmail & ">"
      bolRR = false
      If SendJMail("", strToEmail,"","", strFrName, strFrEmail, strSubj, strMsg,"","",bolRR) = 0 Then
         strEndMsg = "Directory Update Request Sent to " & strToName
      Else
         call Write_Errors
      End If
   End If

   Response.Write "<div align='center'><br><br><b>" & strEndMsg & "</b></div>"

SUB Display_Form
   intSub = 1
   aryData(1,intSub) = "New Email Directory Entry"
   aryData(2,intSub) = aryData(1,intSub)
   aryData(3,intSub) = "H"
   aryData(4,intSub) = "Subj"
   intSub = intSub + 1
   aryData(1,intSub) = "Dean Ammons"
   aryData(2,intSub) = aryData(1,intSub)
   aryData(3,intSub) = "H"
   aryData(4,intSub) = "SN"
   intSub = intSub + 1
   aryData(1,intSub) = "dean@deanammons.com"
   aryData(2,intSub) = aryData(1,intSub)
   aryData(3,intSub) = "H"
   aryData(4,intSub) = "SE"
   intSub = intSub + 1
   aryData(1,intSub) = "First Name"
   aryData(3,intSub) = "T"
   aryData(4,intSub) = "FName"
   intSub = intSub + 1
   aryData(1,intSub) = "Last Name"
   aryData(3,intSub) = "T"
   aryData(4,intSub) = "LName"
   intSub = intSub + 1
   aryData(1,intSub) = "Email"
   aryData(3,intSub) = "T"
   aryData(4,intSub) = "Email"
   intSub = intSub + 1
   aryData(1,intSub) = "Company"
   aryData(3,intSub) = "T"
   aryData(4,intSub) = "ORG"
   intSub = intSub + 1
   aryData(1,intSub) = "Phone"
   aryData(3,intSub) = "T"
   aryData(4,intSub) = "Phone"
   intSub = intSub + 1
   aryData(1,intSub) = "Notes"
   aryData(3,intSub) = "TA"
   aryData(4,intSub) = "Note"

   aryData(0,0) = intSub

   strFldAlign = "R"
   call Get_Data("new_directory_entry.asp","New Directory Entry Details",30)
   Response.End
END SUB

SUB Validate_Input

  '// Validate form entries.

   intErrNbr = 0

   If trim(Request.Form("LName")) = "" Then
      intErrNbr = intErrNbr + 1
      errMsgs.Add intErrNbr, "Last Name is Required"
   End If
   If trim(Request.Form("EMAIL")) = "" and trim(Request.Form("PHONE")) = "" Then
      intErrNbr = intErrNbr + 1
      errMsgs.Add intErrNbr, "Either an Email Address or Telephone Number is Required"
   End If
   If trim(Request.Form("EMAIL")) <> "" Then
      If EmailDuplicate(Request.Form("EMAIL"),0) then
         intErrNbr = intErrNbr + 1
         errMsgs.Add intErrNbr, "Email Address Already In Directory"
      End If
   End If
   If errMsgs.Count > 0 Then
      aryErrorMsgs = errMsgs.Items
   End If
   err.clear
End SUB
%>

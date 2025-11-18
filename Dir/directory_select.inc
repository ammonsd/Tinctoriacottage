
<!--#INCLUDE FILE="../dir/directory_setup.inc" -->
<!--#INCLUDE FILE="../dir/db_global_maint.inc" -->
<!--#INCLUDE FILE="../dir/build_where_sql.inc" -->
<!--#INCLUDE FILE="../dir/validate_email_format.inc" -->
<!--#INCLUDE FILE="../dir/search_prompt.inc" -->
<!--#INCLUDE FILE="../dir/format_security_groups.inc" -->
<!--#INCLUDE FILE="../dir/check_dupe_email.inc" -->

<%
    DIM intRecID, strFName, strLName, strEmail, strKeywords, strContactInfo, strSecGrps, strDateAdded, strDateUpdated
    DIM bolBuildAryOnly, bolAutoSel, strGoToPgm, bolUpdFlds, strCompany, strAddedUserID, strUpdatedUserID

SUB Get_DB_Entries

    DIM intSub, intSub2, intFKey

    Call Setup_Global_Record_Details("")

    If intRecID <> "" and isNumeric(intRecID) Then
       aryDbDtls(0) = int(intRecID)
    End If

    Call Get_Global_Records

    If bolBuildAryOnly Then
       EXIT SUB
    End If

    If bolAutoSel then
       call GoUrl_Code
       Response.Write "<select name='Url' onChange='goUrl(this)'>" & vbCrLf
    Else
       Response.Write "<select name='ID'>" & vbCrLf
    End If

    bolUpdFlds = true

    For intSub = 1 to aryRecData(0,0)
       Call Get_DB_Record(intSub)
       If bolAutoSel then
          Response.Write "<option value='" & strGoToPgm & intRecID & "'>" & strFName & " " & strLName & " (" & strEmail & ")" & "</option>" & vbCrLf
       Else
          Response.Write "<option value='" & intRecID & "'>" & strFName & " " & strLName & "</option>" & vbCrLf
       End If
    Next
    Response.Write "</select>" & vbCrLf

END SUB

SUB Get_DB_Record(ID)

    DIM intSub

    If bolUpdFlds Then
       intSub = ID
    Else
       bolBuildAryOnly = true
       intRecID = ID
       Call Get_DB_Entries
       bolBuildAryOnly = false
       intSub = aryRecData(0,0)
    End If

    If intSub > 0 Then
       intRecID = int(aryRecData(intSub,1))
       strFName = aryRecData(intSub,2)
       strLName = aryRecData(intSub,3)
       strEmail = aryRecData(intSub,4)
       strCompany = aryRecData(intSub,5)
       strSecGrps = UnFormatSecurityGroups(aryRecData(intSub,6))
       strKeywords = aryRecData(intSub,7)
       strContactInfo = aryRecData(intSub,8)
       strDateAdded = aryRecData(intSub,9)
       strAddedUserID = aryRecData(intSub,10)
       strDateUpdated = aryRecData(intSub,11)
       strUpdatedUserID = aryRecData(intSub,12)
    Else
       intRecID = 0
    End If

End SUB

SUB Validate_Entries

   intErrNbr = 0

   strEmail = lcase(strEmail)

   If strFName = "" and strLName = ""  Then
      intErrNbr = intErrNbr + 1
      errMsgs.Add intErrNbr, "Either a First or Last Name Must Be Entered"
   End If
   If strEmail = "" Then
      intErrNbr = intErrNbr + 1
      errMsgs.Add intErrNbr, "Email Address Required"
   ElseIf not ValidateEMailFormat(strEmail) Then
      intErrNbr = intErrNbr + 1
      errMsgs.Add intErrNbr, "Invalid Email Address Format"
   ElseIf strEmail <> Request.Form("OrigEmail") Then
      If EmailDuplicate(strEmail,intRecID) then
         intErrNbr = intErrNbr + 1
         errMsgs.Add intErrNbr, "Email Address Already Registered"
      End If
   End If
   If strSecGrps = ""  Then
      intErrNbr = intErrNbr + 1
      errMsgs.Add intErrNbr, "Security Group Required"
   End If
   If errMsgs.Count > 0 Then
      aryErrorMsgs = errMsgs.Items
   End If
   err.clear

END SUB

SUB Check_Search_Criteria(OPT)

    DIM strHld, intWord, strCde

    Call Setup_Global_Record_Details("")

    If SESSION("KW") <> "" and trim(SESSION("CAT")) = "" and ucase(SESSION("KW")) <> ucase(Application("PGMDIR")) Then
       SESSION("CAT") = SESSION("KW")
       strCatFlds = "COMPANY;#EMAIL"
       bolCatOnly = true
       Call Build_Where_SQL
    End If

    If OPT = "Y" Then
       arySrchOpts(0,0) = "NO"
       arySrchOpts(1,0) = "Name"
       arySrchOpts(2,0) = "Notes"
       If trim(SESSION("CAT")) = "" Then
          arySrchOpts(0,0) = arySrchOpts(0,0) & "C"
          arySrchOpts(3,0) = "Company"
       End If
    End If

    If instr("EWV",SESSION("DEF")) > 0 OR SESSION("SPO") = "Y" Then
       arySpecOpt(1) = "Email"
       arySpecOpt(2) = "Web"
       arySpecOpt(3) = "Display"
       arySpecOpt(0) = 3
       strSpecSecGrp = "VCF"
       Call Verify_Group_Access("N/R","",5)
       If not bolSecError Then
          arySpecOpt(4) = "VCF"
          arySpecOpt(0) = 4
       ElseIf SESSION("DEF") = "V" Then
          SESSION("DEF") = "D"
       End If
       If instr("EWV",SESSION("DEF")) > 0 Then
          bolDefSpecOpt = true
       End If
    End If

    Call Check_Search_Options

    If SESSION("AS") = "Y" Then
       Call Get_Search_Cookies(SESSION("PTYP"))
    End If

    If bolSrchAryOnly Then
       EXIT SUB
    End If

    If CheckSearch Then
       If SESSION("AS") = "Y" Then
          Call Update_Search_Cookies(SESSION("PTYP"))
       End If
       If SESSION("SRCHAREA") = "N" Then
          strSrchFlds = aryRecDtls(2,1) & ";" & aryRecDtls(3,1)
       ElseIf SESSION("SRCHAREA") = "O" Then
          strSrchFlds = aryRecDtls(7,1)
       ElseIf SESSION("SRCHAREA") = "C" Then
          strSrchFlds = aryRecDtls(5,1)
       End If
       If strSrchFlds = "" Then
          strSrchFlds = aryRecDtls(2,1) & ";" & aryRecDtls(3,1) & ";" & aryRecDtls(4,1) & ";" & aryRecDtls(5,1) & ";" & aryRecDtls(7,1) & ";" & aryRecDtls(8,1)
       End If
       Call Build_Where_SQL
    End If

END SUB

SUB Directory_Selection_List(OPT)

    DIM intSub, intSub2, intCntr

    If bolCoSort Then
       aryDbDtls(2) = "5" 'Sort by Company
    Else
       aryDbDtls(2) = "3,2" 'Sort by Last Name, First Name
    End If

    Call Get_DB_Entries

    If bolBuildAryOnly Then
       EXIT SUB
    End If

    If bolAutoSel then
       Response.Write "<select name='Url' onChange='goUrl(this)'>" & vbCrLf
    Else
       Response.Write "<select name='ID'>" & vbCrLf
    End If

    bolUpdFlds = true

    For intSub = 1 to aryRecData(0,0)
       Call Get_DB_Record(intSub)
       If bolAutoSel then
          Response.Write "<option value='" & strGoToPgm & intRecID & "'>" & trim(strFName & " " & strLName) & " (" & strEmail & ")" & "</option>" & vbCrLf
       Else
          Response.Write "<option value='" & intRecID & "'>" & trim(strFName & " " & strLName) & " (" & strEmail & ")" & "</option>" & vbCrLf
       End If
    Next
    Response.Write "</select>" & vbCrLf

END SUB

SUB Setup_Global_Record_Details(TYP)

   Call Database_Setup

   aryDbDtls(0) = 0
   aryDbDtls(1) = APPLICATION("DIRTBL")

   If aryDbDtls(2) = "" Then
       aryDbDtls(2) = "3,2,4" 'Default Sort by Last Name, First Name, email
   End If

   aryRecDtls(1,1) = "ID"
   aryRecDtls(1,2) = CHR(254)
   aryRecDtls(2,1) = "FIRST NAME"
   aryRecDtls(2,2) = strFName
   aryRecDtls(3,1) = "LAST NAME"
   aryRecDtls(3,2) = strLName
   aryRecDtls(4,1) = "EMAIL"
   aryRecDtls(4,2) = strEmail
   aryRecDtls(5,1) = "COMPANY"
   aryRecDtls(5,2) = strCompany
   aryRecDtls(6,1) = "SECURITY GROUPS"
   aryRecDtls(6,2) = FormatSecurityGroups(strSecGrps)
   aryRecDtls(7,1) = "KEYWORDS"
   aryRecDtls(7,2) = strKeywords
   aryRecDtls(8,1) = "CONTACT INFO"
   aryRecDtls(8,2) = strContactInfo
   aryRecDtls(9,1) = "ADDED"
   aryRecDtls(9,3) = "TS"
   aryRecDtls(10,1) = "ADDED USER ID"
   aryRecDtls(10,2) = SESSION("USERID")
   aryRecDtls(11,1) = "UPDATED"
   aryRecDtls(11,3) = "TS"
   aryRecDtls(12,1) = "UPDATE USER ID"
   aryRecDtls(12,2) = SESSION("USERID")

   If TYP = "A" Then
      aryRecDtls(11,2) = CHR(254)
      aryRecDtls(12,2) = CHR(254)
   Else
      aryRecDtls(9,2) = CHR(254)
      aryRecDtls(10,2) = CHR(254)
  End If

END SUB
%>



<!--#INCLUDE FILE="../dir/member_setup.inc" -->
<!--#INCLUDE FILE="../dir/db_global_maint.inc" -->
<!--#INCLUDE FILE="../dir/build_where_sql.inc" -->
<!--#INCLUDE FILE="../dir/validate_email_format.inc" -->
<!--#INCLUDE FILE="../dir/search_prompt.inc" -->
<!--#INCLUDE FILE="../dir/state_country_select.inc" -->
<!--#INCLUDE FILE="../dir/check_dupe_email.inc" -->

<%
    DIM strFName, strLName, strEmail, strStreet, strCity, strState, strZip, strCountry, strComments
    DIM bolBuildAryOnly, bolAutoSel, intRecID, strGoToPgm, bolUpdFlds
    DIM bolMailingList, strDateAdded, strDateUpdated, strCompany, strAddress2, strPhoneNbr, strKeyWords

SUB Get_DB_Entries

    DIM intSub, intSub2, intSub3, intFKey

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
          Response.Write "<option value='" & strGoToPgm & intRecID & "'>" & strFName & " " & strLName & "</option>" & vbCrLf
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
       strStreet = aryRecData(intSub,6)
       strAddress2 = aryRecData(intSub,7)
       strCity = aryRecData(intSub,8)
       strState = aryRecData(intSub,9)
       strZip = aryRecData(intSub,10)
       strCountry = aryRecData(intSub,11)
       strPhoneNbr = aryRecData(intSub,12)
       bolMailingList = aryRecData(intSub,13)
       strKeyWords = aryRecData(intSub,14)
       strDateAdded = aryRecData(intSub,15)
       strDateUpdated = aryRecData(intSub,17)
    Else
       intRecID = 0
    End If

End SUB

FUNCTION GetMbrID(EMAIL)

    GetMbrID = 0

    If EMAIL = "" or APPLICATION("MBRTBL") = "" Then
       EXIT FUNCTION
    End If

    Call Database_Setup

    Set objRecordSet = Server.CreateObject("ADODB.Recordset")
    objRecordSet.ActiveConnection = strDbConn

    objRecordSet.Source = "SELECT [ID] FROM " & APPLICATION("MBRTBL") & " WHERE [eMail] = '" & lcase(EMAIL) & "'"
    objRecordSet.LockType = adLockReadOnly
    objRecordSet.Open

    If Not objRecordSet.EOF Then
       objRecordSet.MoveFirst
       GetMbrID = objRecordSet.Fields("ID")
    End If

    objRecordSet.Close
    Set objRecordSet = Nothing

END FUNCTION

SUB Validate_Entries

   intErrNbr = 0

   strEmail = lcase(strEmail)

   If strFName = "" and strLName = ""  Then
      intErrNbr = intErrNbr + 1
      errMsgs.Add intErrNbr, "First or Last Name Required"
   End If
   If strEmail = "" Then
      intErrNbr = intErrNbr + 1
      errMsgs.Add intErrNbr, "E-Mail Address Required"
   ElseIf not ValidateEMailFormat(strEmail) Then
      intErrNbr = intErrNbr + 1
      errMsgs.Add intErrNbr, "Invalid E-Mail Address Format"
   ElseIf strEmail <> Request.Form("OrigEmail") Then
      If GetMbrID(strEmail) <> intRecID then
         intErrNbr = intErrNbr + 1
         errMsgs.Add intErrNbr, "Email Address Already Registered"
      End If
   End If
   If trim(strZip) <> "" Then
      If len(strZip) < 5 or len(strZip) > 10 or CheckNumeric(strZip,"-") = false Then
         intErrNbr = intErrNbr + 1
         errMsgs.Add intErrNbr, "Invalid Zip Code Format"
      ElseIf len(strZip) > 5 Then
         If mid(strZip,6,1) <> "-" or len(strZip) <> 10 Then
            intErrNbr = intErrNbr + 1
            errMsgs.Add intErrNbr, "Invalid Zip+4 Format (nnnnn-nnnn)"
         End If
      End If
   End If
   If trim(Request.Form("UserName")) <> "" Then
      If trim(Request.Form("Password")) <> trim(Request.Form("Password2")) Then
         intErrNbr = intErrNbr + 1
         errMsgs.Add intErrNbr, "Passwords Don't Match - Please Re-Enter"
      ElseIf len(trim(Request.Form("Password"))) < 6 Then
         intErrNbr = intErrNbr + 1
         errMsgs.Add intErrNbr, "New Password Must Be at Least 6 Characters"
      End If
   End If
   If strCountry = ""  Then
      intErrNbr = intErrNbr + 1
      errMsgs.Add intErrNbr, "Country Required"
   End If
   If errMsgs.Count > 0 Then
      aryErrorMsgs = errMsgs.Items
   End If
   err.clear

END SUB

SUB Check_Search_Criteria

    DIM strHld, intSub, intWord, strCde

    arySrchOpts(0,0) = "SN"
    arySrchOpts(1,0) = "State"
    arySrchOpts(2,0) = "Notes"

    Call Check_Search_Options

    Call Setup_Global_Record_Details("")

    If SESSION("AS") = "Y" and SESSION("MC") = "" Then
       Call Get_Search_Cookies(SESSION("PTYP"))
    End If

    If bolSrchAryOnly Then
       EXIT SUB
    End If

    If CheckSearch  Then
       If SESSION("SRCHAREA") = "S" Then
          If len(SESSION("SRCHKW")) > 2 Then
             bolSelAryOnly = true
             Call Build_States_Select_List("")
             SESSION("SRCHKW") = trim(RemoveLineBreaks(SESSION("SRCHKW")))
             bolQuoteIsPhrase = true
             intWord = 1
             strCde = GetWords(SESSION("SRCHKW"),intWord,1)
             strHld = ""
             Do while strCde <> ""
                If len(strCde) > 2 and ucase(strCde) <> "AND" Then
                   For intSub = intBlnkEntry+1 to arySelData(0,0)
                      If ucase(arySelData(intSub,2)) = ucase(strCde) Then
                         strCde = arySelData(intSub,1)
                         EXIT FOR
                      End If
                   Next
                End If
                If strHld = "" Then
                   strHld = strCde
                Else
                   strHld = strHld & " " & strCde
                End If
                intWord = intWord + intRetWords
                strCde = GetWords(SESSION("SRCHKW"),intWord,1)
             Loop
             SESSION("SRCHKW") = strHld
          End If
          bolMatchWord = true
          strSrchFlds = aryRecDtls(9,1)
       ElseIf SESSION("SRCHAREA") = "N" Then
          strSrchFlds = aryRecDtls(14,1)
       End If

       If SESSION("AS") = "Y" Then
          Call Update_Search_Cookies(SESSION("PTYP"))
       End If

       If strSrchFlds = "" Then
          strSrchFlds = aryRecDtls(2,1) & ";" & aryRecDtls(3,1) & ";" & aryRecDtls(4,1) & ";" & aryRecDtls(5,1) & ";" & aryRecDtls(14,1)
       End If
       Call Build_Where_SQL
    End If

END SUB

SUB Setup_Global_Record_Details(TYP)

   Call Database_Setup

   aryDbDtls(0) = 0
   aryDbDtls(1) = APPLICATION("MBRTBL")

   If aryDbDtls(2) = "" Then
       aryDbDtls(2) = "3,2" 'Default Sort by Last Name, First Name
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
   aryRecDtls(6,1) = "STREET"
   aryRecDtls(6,2) = strStreet
   aryRecDtls(7,1) = "ADDRESS2"
   aryRecDtls(7,2) = strAddress2
   aryRecDtls(8,1) = "CITY"
   aryRecDtls(8,2) = strCity
   aryRecDtls(9,1) = "STATE"
   aryRecDtls(9,2) = strState
   aryRecDtls(10,1) = "ZIP"
   aryRecDtls(10,2) = strZip
   aryRecDtls(11,1) = "COUNTRY"
   aryRecDtls(11,2) = strCountry
   aryRecDtls(12,1) = "PHONE NUMBER"
   aryRecDtls(12,2) = strPhoneNbr
   aryRecDtls(13,1) = "MAILING LIST"
   aryRecDtls(13,2) = bolMailingList
   aryRecDtls(14,1) = "KEYWORDS"
   aryRecDtls(14,2) = strKeyWords
   aryRecDtls(15,1) = "ADDED"
   aryRecDtls(15,3) = "TS"
   aryRecDtls(16,1) = "ADDED USER ID"
   aryRecDtls(16,2) = SESSION("USERID")
   aryRecDtls(17,1) = "UPDATED"
   aryRecDtls(17,3) = "TS"
   aryRecDtls(18,1) = "UPDATE USER ID"
   aryRecDtls(18,2) = SESSION("USERID")

   If TYP = "A" Then
      aryRecDtls(17,2) = CHR(254)
      aryRecDtls(18,2) = CHR(254)
   Else
      aryRecDtls(15,2) = CHR(254)
      aryRecDtls(16,2) = CHR(254)
  End If

END SUB

%>


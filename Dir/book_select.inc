
<!--#INCLUDE FILE="../dir/book_setup.inc" -->
<!--#INCLUDE FILE="../dir/db_global_maint.inc" -->
<!--#INCLUDE FILE="../includes/setup_book_cat.inc" -->
<!--#INCLUDE FILE="../dir/build_where_sql.inc" -->
<!--#INCLUDE FILE="../dir/search_prompt.inc" -->

<%
    DIM bolBuildAryOnly, bolAutoSel, aryCategory(40), strGoToPgm, bolUpdFlds, bolNoFormat, strLang
    DIM intRecID, strTitle, strPicFile, strDetails, strAuthor, strCategory, bolInactive, strISBN, strCrDate, strPicFileSm

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
          Response.Write "<option value='" & strGoToPgm & intRecID & "'>" & strTitle & "</option>" & vbCrLf
       Else
          Response.Write "<option value='" & intRecID & "'>" & strTitle & "</option>" & vbCrLf
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
       strCategory = aryRecData(intSub,2)
       strTitle = aryRecData(intSub,3)
       strAuthor = aryRecData(intSub,4)
       strPicFile = lcase(aryRecData(intSub,5))
       If strPicFile <> "" Then
          strPicFileSm = left(strPicFile,instr(strPicFile,".")-1) & "_t" & mid(strPicFile,instr(strPicFile,"."))
       End If
       strDetails = aryRecData(intSub,6)
       strISBN = aryRecData(intSub,7)
       If strISBN <> "" and not bolNoFormat and instr(strISBN,"-") = 0 Then
          strISBN = left(strISBN,1) & "-" & mid(strISBN,2,5) & "-" & mid(strISBN,7,3) & "-" & mid(strISBN,10,1)
       End If
       strCrDate = aryRecData(intSub,8)
       If strCrDate = "0" Then
          strCrDate = ""
       End If
       strLang = aryRecData(intSub,9)
       bolInactive = aryRecData(intSub,10)
    Else
       intRecID = 0
    End If

End SUB

SUB Build_Category_Selection(CURCAT)

    DIM intSub, strCat

    call Build_Book_Category_Array

    If bolBuildAryOnly Then
       EXIT SUB
    End If

    For intSub = 1 to aryCategory(0)
       If aryCategory(intSub) = CURCAT Then
           strCat = CURCAT
           EXIT FOR
       End If
    Next

    Response.Write "<select name='Category'>" & vbCrLf
    If strCat <> "" Then
       Response.Write "<option selected value='" & strCat & "'>" & strCat & vbCrLf
    End IF

    For intSub = 1 to aryCategory(0)
       If aryCategory(intSub) <> strCat Then
          Response.Write "<option value='" & aryCategory(intSub) & "'>" & aryCategory(intSub) & "</option>" & vbCrLf
       End If
    Next
    Response.Write "</select>" & vbCrLf


END SUB

SUB Validate_Entries

   If Request.Form("Status") = "I" Then
      bolInactive = true
   Else
      bolInactive = false
   End If

   If strTitle = "" Then
      intErrNbr = intErrNbr + 1
      errMsgs.Add intErrNbr, "Title Required"
   End If
   If strAuthor = "" Then
      intErrNbr = intErrNbr + 1
      errMsgs.Add intErrNbr, "Author Required"
   End If
   If strCategory = "" then
      intErrNbr = intErrNbr + 1
      errMsgs.Add intErrNbr, "Category Required"
   End If
   If strCrDate <> "" Then
      If not isNumeric(strCrDate) Then
         intErrNbr = intErrNbr + 1
         errMsgs.Add intErrNbr, "Copyright Year Must Be Numeric (YYYY)"
      ElseIf int(strCrDate) < 1000 Then
         intErrNbr = intErrNbr + 1
         errMsgs.Add intErrNbr, "Format of Copyright Year Invalid (YYYY)"
      End If
   Else
      strCrDate = "0"
   End If
   If strISBN <> "" and strISBN = "" Then
      If (instr(strISBN,"-") > 0 and len(strISBN) <> 13) or (instr(strISBN,"-") = 0 and len(strISBN) <> 10) Then
            intErrNbr = intErrNbr + 1
            errMsgs.Add intErrNbr, "ISBN must be 10 characters (Unformatted) or 13 characters (Formatted)"
      End If
   End If

   If strLang = "" Then
      strLang = strDefBookLang
   End If
END SUB

SUB Check_Search_Criteria(CHKSTATUS)

    Call Setup_Global_Record_Details("")

    arySrchOpts(0,0) = "TAD"
    arySrchOpts(1,0) = "Title"
    arySrchOpts(2,0) = "Author"
    arySrchOpts(3,0) = "Description"
    If trim(SESSION("CAT")) = "" and Request.QueryString("CAT") = "" Then
       arySrchOpts(0,0) = arySrchOpts(0,0) & "C"
       arySrchOpts(4,0) = "Category"
       arySrchOpts(4,1) = " onClick=" & CHR(34) & "document.location.href ='select_book_cat.asp?P=#PGM#'" & CHR(34)
    End If

    If CHKSTATUS <> "N" Then
       aryStatusOpts(0) = "AI"
       aryStatusOpts(1) = "Active"
       aryStatusOpts(2) = "Inactive"
    End If

    Call Check_Search_Options

    If SESSION("AS") = "Y" and trim(SESSION("CAT")) = "" Then
       Call Get_Search_Cookies(SESSION("PTYP"))
    End If

    If Request.QueryString("CAT") <> "" Then
       SESSION("CAT") = ucase(Request.QueryString("CAT"))
       strCatFlds = "Category"
       If SESSION("CAT") = "-" Then
          Session.Contents.Remove("CAT")
          Session.Contents.Remove("RECCNT")
       End If
       bolCatOnly = true
       Call Build_Where_SQL
    End If

    If bolSrchAryOnly Then
       EXIT SUB
    End If

    If CheckSearch Then
       If SESSION("AS") = "Y" and trim(SESSION("CAT")) = "" Then
          Call Update_Search_Cookies(SESSION("PTYP"))
       End If
       If SESSION("SRCHAREA") = "T" Then
          strSrchFlds = aryRecDtls(3,1)
       ElseIf SESSION("SRCHAREA") = "A" Then
          strSrchFlds = aryRecDtls(4,1)
       ElseIf SESSION("SRCHAREA") = "D" Then
          strSrchFlds = aryRecDtls(6,1)
       ElseIf SESSION("SRCHAREA") = "C" Then
          bolMatchWord = true
          strSrchFlds = aryRecDtls(2,1)
       End If

       If strSrchFlds = "" Then
          strSrchFlds = aryRecDtls(3,1) & ";" & aryRecDtls(4,1) & ";" & aryRecDtls(6,1)
       End If
       Call Build_Where_SQL
    End If

END SUB

SUB Build_Book_Details

   strDetails = RemoveLineBreaks(strDetails)
   If not bolNoFormat Then
      strDetails = REPLACE(strDetails,CHR(34),"`")
      strDetails = REPLACE(strDetails,"'","`")
   End If

   If strISBN <> "" Then
      strDetails = strDetails & "&nbsp;&nbsp; ISBN: " & strISBN
   End If
   If strCrDate <> "" Then
      strDetails = strDetails & "&nbsp;&nbsp; Published: " & strCrDate
   End If
   If strLang <> ""  Then
      strDetails = strDetails & "&nbsp;&nbsp; Language: " & strLang
   End If

END SUB

SUB Setup_Global_Record_Details(TYP)

   DIM intSub, intAddSub, intUpdSub

   Call Database_Setup

   aryDbDtls(0) = 0
   aryDbDtls(1) = APPLICATION("BOOKTBL")

   If aryDbDtls(2) = "" Then
       aryDbDtls(2) = "3,4" 'Default Sort by Title, Author
   End If

   aryRecDtls(1,1) = "ID"
   aryRecDtls(1,2) = CHR(254)
   aryRecDtls(2,1) = "Category"
   aryRecDtls(2,2) = strCategory
   aryRecDtls(3,1) = "Title"
   aryRecDtls(3,2) = strTitle
   aryRecDtls(4,1) = "Author"
   aryRecDtls(4,2) = strAuthor
   aryRecDtls(5,1) = "Picture File"
   aryRecDtls(5,2) = strPicFile
   aryRecDtls(6,1) = "Details"
   aryRecDtls(6,2) = strDetails
   aryRecDtls(7,1) = "ISBN"
   aryRecDtls(7,2) = strISBN
   aryRecDtls(8,1) = "Copyright"
   aryRecDtls(8,2) = strCrDate
   aryRecDtls(8,3) = "I"
   aryRecDtls(9,1) = "Language"
   aryRecDtls(9,2) = strLang
   intSub = 10
   aryRecDtls(intSub,1) = "Inactive"
   aryRecDtls(intSub,2) = bolInactive
   intSub = intSub + 1
   aryRecDtls(intSub,1) = "ADDED"
   aryRecDtls(intSub,3) = "TS"
   intAddSub = intSub
   intSub = intSub + 1
   aryRecDtls(intSub,1) = "ADDED USER ID"
   aryRecDtls(intSub,2) = SESSION("USERID")
   intSub = intSub + 1
   aryRecDtls(intSub,1) = "UPDATED"
   aryRecDtls(intSub,3) = "TS"
   intUpdSub = intSub
   intSub = intSub + 1
   aryRecDtls(intSub,1) = "UPDATE USER ID"
   aryRecDtls(intSub,2) = SESSION("USERID")

   If TYP = "A" Then
      aryRecDtls(intUpdSub,2) = CHR(254)
      aryRecDtls(intUpdSub+1,2) = CHR(254)
   Else
      aryRecDtls(intAddSub,2) = CHR(254)
      aryRecDtls(intAddSub+1,2) = CHR(254)
  End If

END SUB

%>


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
<!--#INCLUDE FILE="../dir/get_data.inc" -->
<!--#INCLUDE FILE="../dir/display_edit_Errors.inc" -->

<%

   DIM intSub, bolRecFnd, strDbAction, intErrNbr, strTemp, strID, strMaintType, strHld, strTabChr, intCnt
   DIM strReqFn, intLoc, strPrevCategory

   strMaintType = "A"
   strDbAction = "A"
   strTabChr = CHR(9)

   Call System_Setup("NONE")
   Call Logon_Check(GetCurPath("")&"book_batch_processing.asp",3,strLogonGrp)

   strGetData = SESSION("GETDATA")
   Session.Contents.Remove("GETDATA")

   strReqFn = Request.QueryString("FN")

   If Request.QueryString("UPD") = "Y" and strReqFn <> "" Then
      SESSION("PGM") = GetCurPath("")&"book_batch_processing.asp?FN=" & strReqFn
      If Request.QueryString("M") <> "" Then
         SESSION("PGM") = SESSION("PGM") & "&M=" & Request.QueryString("M")
      End If
      SESSION("FN") = strReqFn
      Response.ReDirect(GetCurPath("")&"maint_text_file.asp?E=Y&Msg=** Paste Batch Book Entries Over This Line **")
   ElseIf SESSION("PGM") <> "" Then
      SESSION("PGM") = ""
      SESSION("FN") = ""
   End If

   Call Database_Setup

   Set errMsgs = CreateObject("Scripting.Dictionary")

   If strReqFn = "" Then
      strReqFn = Request.Form("FN")
   End If
   If strReqFn = "" Then
      aryData(1,1) = "Name of File with Book Entries"
      aryData(2,1) = GetCookie("BOOKBATCH","REQFN")
      aryData(3,1) = ""
      aryData(4,1) = "FN"
      aryData(0,0) = 1
      If strGetData <> "Y" Then
         call Get_Data("book_batch_processing.asp","Batch Book Entries",30)
      End If
      Response.End
   End If

   If CheckFileExist("/" & APPLICATION("FILEDIR") & "/" & strReqFn) <> 0 Then
      intErrNbr = intErrNbr + 1
      errMsgs.Add intErrNbr, "The Input File " & CHR(34) & strReqFn & CHR(34) & " was not Found in " & CHR(34) & "/" & APPLICATION("FILEDIR") & "/" & CHR(34)
   End If

   If errMsgs.Count > 0 Then
      aryErrorMsgs = errMsgs.Items
      err.clear
      call Display_Errors
   End If

   Call AddCookie("BOOKBATCH","REQFN",strReqFn,365)
   Call Read_Text_File(strReqFn,Application("FILEDIR"))

   Set errMsgs = CreateObject("Scripting.Dictionary")
   intCnt = 0
   strPrevCategory = "None"

   For intSub = 2 to aryTxtRecs(0) 'Skip heading column
      If instr(aryTxtRecs(intSub),strTabChr) > 0 Then
         strCategory = ""
         strLang = ""
         strHld = aryTxtRecs(intSub) & strTabChr
         intLoc = instr(strHld,strTabChr)
         If intLoc > 0 Then
            strTitle = trim(left(strHld,intLoc-1))
            strHld = mid(strHld,intLoc+1)
           If left(strTitle,1) = CHR(34) and right(strTitle,1) = CHR(34) Then
              strTitle = mid(strTitle,2)
              strTitle = left(strTitle,len(strTitle)-1)
           End If
         Else
            strTitle = ""
         End If
         intLoc = instr(strHld,strTabChr)
         If intLoc > 0 Then
           strAuthor = trim(left(strHld,intLoc-1))
           strHld = mid(strHld,intLoc+1)
           If left(strAuthor,1) = CHR(34) and right(strAuthor,1) = CHR(34) Then
              strAuthor = mid(strAuthor,2)
              strAuthor = left(strAuthor,len(strAuthor)-1)
           End If
         Else
            strAuthor = ""
         End If
         intLoc = instr(strHld,strTabChr)
         If intLoc > 0 Then
           strDetails = trim(left(strHld,intLoc-1))
           strHld = mid(strHld,intLoc+1)
           If left(strDetails,1) = CHR(34) and right(strDetails,1) = CHR(34) Then
              strDetails = mid(strDetails,2)
              strDetails = left(strDetails,len(strDetails)-1)
           End If
         Else
            strDetails = ""
         End If
         intLoc = instr(strHld,strTabChr)
         If intLoc > 0 Then
           strCategory = trim(left(strHld,intLoc-1))
           If trim(strCategory) = "" Then
              strCategory = "None"
           End If
           strHld = mid(strHld,intLoc+1)
         Else
            strCategory = ""
         End If
         If strTitle <> "" and strAuthor <> "" Then
            intLoc = instr(strHld,strTabChr)
            If intLoc > 0 Then
              strPicFile = trim(left(strHld,intLoc-1))
              strHld = mid(strHld,intLoc+1)
            Else
               strPicFile = ""
            End If
            intLoc = instr(strHld,strTabChr)
            If intLoc > 0 Then
              strISBN = trim(left(strHld,intLoc-1))
              strHld = mid(strHld,intLoc+1)
            Else
               strISBN = ""
            End If
            intLoc = instr(strHld,strTabChr)
            If intLoc > 0 Then
              strCrDate = trim(left(strHld,intLoc-1))
              strHld = mid(strHld,intLoc+1)
            Else
               strCrDate = ""
            End If
            intLoc = instr(strHld,strTabChr)
            If intLoc > 0 Then
              strLang = trim(left(strHld,intLoc-1))
              strHld = mid(strHld,intLoc+1)
            Else
               strLang = ""
            End If
            bolInactive = false
            If trim(strCategory) = "" Then
               strCategory = strPrevCategory
            End If
            strPrevCategory = strCategory
            If trim(strLang) = "" Then
               strLang = strCategory
            End If

            intCnt = intCnt + 1
            Call Validate_Entries
            If errMsgs.Count > 0 Then
               strDisplayMsg = "Errors Found in Record #" & intCnt
               If intCnt > 1 Then
                  If intCnt = 2 Then
                     strDisplayMsg = strDisplayMsg & " - Record# 1 was added to the database"
                  Else
                     strDisplayMsg = strDisplayMsg & " - Records 1-" & (intCnt-1) & " were added to the database"
                  End If
               End If
               aryErrorMsgs = errMsgs.Items
               err.clear
               call Display_Errors
            Else
               Call Add_Record
               If errMsgs.Count > 0 Then
                  aryErrorMsgs = errMsgs.Items
                  err.clear
                  call Process_Errors
               ElseIf Err.Number <> 0 and Err.Number <> 53 Then
                 call Write_Errors
               End If
            End If
         End If
      End If
   Next

   aryTxtRecs(0) = 0 'Delete input file
   Call Write_Text_File(strReqFn,Application("FILEDIR"))

   bolNoEnd = true

   call Display_Ending_Msg(intCnt & " New Book Entries Added to Database" ,"")

   If Request.QueryString("M") <> "" Then
       Response.Write "<div align='center'>" & vbCrLf
       Response.Write "<br><br><input type=" & CHR(34) & "button" & CHR(34) & " value=" & CHR(34) & "Menu" & CHR(34) & " onClick=" & CHR(34) & "document.location.href = '" & Request.QueryString("M") & "'" & CHR(34) & ">" & vbCrLf
       Response.Write "</div'>" & vbCrLf
   End If

SUB Add_Record

   Call Setup_Global_Record_Details("A")
   Call AddGlobalRecord

END SUB

%>

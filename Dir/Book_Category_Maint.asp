<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/maint_webpage.inc" -->
<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->
<!--#INCLUDE FILE="../dir/book_setup.inc" -->

<%
    DIM intSub, strReqFn, intCntr, strPgm, strHld, intRows, strIncFn, X, X2, intSub2, aryInc(40), strMenu

    strPgm = "book_category_maint.asp"

    Call System_Setup("NONE")
    Call Logon_Check(GetCurPath("")&strPgm,3,strLogonGrp)

    strMenu = Request.QueryString("M")
    If strMenu = "" Then
       strMenu = Request.Form("M")
    End If

    strReqFn = "bookcat.dat"
    strIncFn = "setup_book_cat.inc"

    If Request.Form("UpdFile") = "Y" Then
       intSub2 = 0
       For intSub = 1 to 40
          X = "Cat" & intSub
          X2 = "OCat" & intSub
          If Request.Form(X) <> "" Then
             If Request.Form(X) <> Request.Form(X2) Then
                Call Change_Category(Request.Form(X2),Request.Form(X))
             End If
             intSub2 = intSub2 + 1
             aryTxtRecs(intSub2) = Request.Form(X)
             aryInc(intSub2) = Request.Form(X)
          End If
       Next
       If intSub2 = 0 Then
          intSub2 = 1
          aryTxtRecs(intSub2) = "None"
       End If
       aryTxtRecs(0) = intSub2
       aryInc(0) = intSub2
       call Write_Text_File(strReqFn,"CGI-BIN")

       intSub = 1
       aryTxtRecs(intSub) = "<%SUB Build_Book_Category_Array"
       For intSub2 = 1 to aryInc(0)
          intSub = intSub + 1
          aryTxtRecs(intSub) = "aryCategory(" & intSub2 & ") = " & CHR(34) &  aryInc(intSub2) & CHR(34)
       Next
       intSub = intSub + 1
       aryTxtRecs(intSub) = "aryCategory(0) = " & aryInc(0)
       intSub = intSub + 1
       aryTxtRecs(intSub) = "END SUB" & CHR(37) & ">"
       aryTxtRecs(0) = intSub

       call Write_Text_File(strIncFn,Application("INCLDIR"))

       Response.Write "<div align='center'>" & vbCrLf
       Response.Write "<br><br><b><br><br><b><font color='red'>Book Categories Updated</font></b>" & vbCrLf
       If strMenu <> "" Then
           Response.Write "<br><br><input type=" & CHR(34) & "button" & CHR(34) & " value=" & CHR(34) & "Return to Menu" & CHR(34) & " onClick=" & CHR(34) & "document.location.href = '" & strMenu & "'" & CHR(34) & ">" & vbCrLf
       End If
       Response.Write "</div'>" & vbCrLf
       Response.End
    End If

    call Read_Text_File(strReqFn,"CGI-BIN")

    strMaint = "Y"
    strTblClr = "#b0c4de"
    strFSize = "+1"
    call Setup_Web_Page("Update Book Categories",8)

    Response.Write "<form action='book_category_maint.asp' method='post'>" & vbCrLf
    Response.Write "<input type='hidden' name='UpdFile' value='Y'>" & vbCrLf
    If strMenu <> "" Then
       Response.Write "<input type='hidden' name='M' value='" & strMenu &"'>" & vbCrLf
    End If

    For intSub = 1 to 10
       Response.Write "<tr><td>" & vbCrLf
       Response.Write "#" & intSub & vbCrLf
       Response.Write "</td><td>" & vbCrLf
       Response.Write "<input type='text' name='Cat" & intSub & "' size='20' maxlength='50' value='" & aryTxtRecs(intSub) & "'>" & vbCrLf
       Response.Write "<input type='hidden' name='OCat" & intSub & "' value='" & aryTxtRecs(intSub) & "'>" & vbCrLf
       Response.Write "</td><td>" & vbCrLf
       Response.Write "#" & intSub+10 & vbCrLf
       Response.Write "</td><td>" & vbCrLf
       Response.Write "<input type='text' name='Cat" & intSub+10 & "' size='20' maxlength='50' value='" & aryTxtRecs(intSub+10) & "'>" & vbCrLf
       Response.Write "<input type='hidden' name='OCat" & intSub+10 & "' value='" & aryTxtRecs(intSub+10) & "'>" & vbCrLf
       Response.Write "</td><td>" & vbCrLf
       Response.Write "#" & intSub+20 & vbCrLf
       Response.Write "</td><td>" & vbCrLf
       Response.Write "<input type='text' name='Cat" & intSub+20 & "' size='20' maxlength='50' value='" & aryTxtRecs(intSub+20) & "'>" & vbCrLf
       Response.Write "<input type='hidden' name='OCat" & intSub+20 & "' value='" & aryTxtRecs(intSub+20) & "'>" & vbCrLf
       Response.Write "</td><td>" & vbCrLf
       Response.Write "#" & intSub+30 & vbCrLf
       Response.Write "</td><td>" & vbCrLf
       Response.Write "<input type='text' name='Cat" & intSub+30 & "' size='20' maxlength='50' value='" & aryTxtRecs(intSub+30) & "'>" & vbCrLf
       Response.Write "<input type='hidden' name='OCat" & intSub+30 & "' value='" & aryTxtRecs(intSub+30) & "'>" & vbCrLf
       Response.Write "</td></tr>" & vbCrLf
    Next

    Response.Write "<tr><td colspan='8'>" & vbCrLf
    Response.Write "<hr width='100%' size='5' noshade>" & vbCrLf
    Response.Write "</td></tr>" & vbCrLf
    Response.Write "<tr><td align='center' colspan='8'>" & vbCrLf
    Response.Write "<input type='SUBMIT' name='Update' value='Update'>" & vbCrLf
    Response.Write "</form>" & vbCrLf
    Response.Write "</td></tr>" & vbCrLf
    Call Maintenance_Links(8)
    Response.Write "</table>" & vbCrLf
    Response.Write "<div align='center'>" & vbCrLf
    Response.Write "<font color='red' size=-1>Leave Category Blank to Remove Entry</font>" & vbCrLf
    Response.Write "</div></form>" & vbCrLf

SUB Change_Category(OldCat,NewCat)

   DIM strSQL

   Call Database_Setup

   Set objRecordSet = Server.CreateObject("ADODB.Recordset")
   objRecordSet.LockType = adLockPessimistic
   objRecordSet.CursorLocation = adUseClient
   objRecordSet.CursorType = adOpenStatic
   strSQL = "UPDATE " & APPLICATION("BOOKTBL") & " SET [CATEGORY] = '" & NewCat & "' WHERE [TYPE] = 'BOOK' AND [CATEGORY] = '" & OldCat & "'"
   objRecordSet.Open strSQL, strDbConn
   Set objRecordSet = Nothing

END SUB
%>


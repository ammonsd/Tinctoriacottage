<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/maint_webpage.inc" -->
<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->
<!--#INCLUDE FILE="../dir/photo_details_setup.inc" -->
<!--#INCLUDE FILE="../dir/db_gen_select.inc" -->
<!--#INCLUDE FILE="../dir/db_gen_maint.inc" -->

<%
    DIM intSub, intSub2, intSub3, intCntr, strPgm, strHld, X, bolUpOK, strTitle, intSeq, strMenu
    DIM strPID, intPID, intDID, intCID, strWebName, strPicFile, strDesc, strCaption, strWebSite

    strPgm = "photo_details_maint.asp"
    bolNoSrch = true

    strWebName = Request.QueryString("WN")
    If strWebName = "" Then
       strWebName = Request.Form("WN")
    End If
    strTitle = Request.QueryString("TI")
    If strTitle = "" Then
       strTitle = Request.Form("TI")
    End If
    If strTitle = "" Then
       strTitle = strWebName
    End If
    strWebSite = Request.QueryString("WS")
    If strWebSite = "" Then
       strWebSite = Request.Form("WS")
    End If
    strMenu = Request.QueryString("M")
    If strMenu = "" Then
       strMenu = Request.Form("M")
    End If

    Call System_Setup("NONE")
    Call Logon_Check(GetCurPath("")&strPgm,3,strLogonGrp)
    Call Database_Setup

    If Request.Form("UpdFile") = "Y" Then
       intSub2 = 0
       intLastSeq = 0
       For intSub = 1 to SESSION("CNT")
          X = "Pho" & intSub
          strPicFile = lcase(Request.Form(X))
          X = "PID" & intSub
          strPID = Request.Form(X)
          X = "Dsc" & intSub
          strDesc = Request.Form(X)
          X = "DID" & intSub
          If Request.Form(X)  = "" or not isNumeric(Request.Form(X)) Then
             intDID = 0
          Else
             intDID = int(Request.Form(X))
          End If
          X = "CAP" & intSub
          strCaption = Request.Form(X)
          X = "CID" & intSub
          If Request.Form(X)  = "" or not isNumeric(Request.Form(X)) Then
             intCID = 0
          Else
             intCID = int(Request.Form(X))
          End If
          X = "SEQ" & intSub
          If Request.Form(X)  = "" or not isNumeric(Request.Form(X)) Then
             intSeq = intLastSeq + 1
          Else
             intSeq = int(Request.Form(X))
          End If
          If intSeq > intLastSeq Then
             intLastSeq = intSeq
          End If
          If strPicFile <> "" or strPID <> "" Then
             If strPID = "" or not isNumeric(strPID) Then
                intPID = 0
             Else
                intPID = int(strPID)
             End If
             If intPID = 0 Then
                intFKey = AddGenTextRec("PHOTO",ucase(strWebName),intSeq,strPicFile,false,"","")
                If strCaption <> "" Then
                   Call AddGenTextRec("PHOTO",ucase(strWebName),intSeq,strCaption,false,intFKey,"")
                End If
                If strDesc <> "" Then
                   Call AddGenMemoRec("PHOTO",ucase(strWebName),0,strDesc,intFKey)
                End If
             ElseIf intPID > 0 and (intDID = 0 or intCID = 0) Then
                If intCID = 0 and strCaption <> "" Then
                   Call AddGenTextRec("PHOTO",ucase(strWebName),intSeq,strCaption,false,intPID,"")
                End If
                If intDID = 0 and strDesc <> "" Then
                   Call AddGenMemoRec("PHOTO",ucase(strWebName),0,strDesc,intPID)
                End If
             Else
                If strPicFile = "" Then
                   strPicFile = ""
                   Call CascadeDelGenTextRec(intPID)
                Else
                   If strDesc = "" and intDID > 0 Then
                     Call DelGenMemoRec(intDID)
                     intDID = 0
                   End If
                   If strCaption = "" and intCID > 0 Then
                     Call DelGenTextRec(intCID)
                     intCID = 0
                   End If
                End If
             End If
             If intPID > 0 and strPicFile <> "" Then
                Call UpdGenTextRec(intPID,ucase(strWebName),intSeq,strPicFile,false)
                If intDID > 0 Then
                   Call UpdGenMemoRec(intDID,strDesc)
                End If
                If intCID > 0 Then
                   Call UpdGenTextRec(intCID,ucase(strWebName),intSeq,strCaption,false)
                End If
             End If
          End If
       Next

       SESSION("CNT") = ""
       'Build SSI File
       Call Get_Gen_Text_Recs("PHOTO",ucase(strWebName))
       Call Get_Gen_Text2_Recs(ucase(strWebName))

       intCntr = 1
       aryTxtRecs(intCntr) = "<br>"
       For intSub = 1 to aryGenText(0,0)
          strPicFile = aryGenText(intSub,5)
          If strPicFile <> "" Then
             strCaption = GetPhotoCaption(aryGenText(intSub,1),ucase(strWebName))
             Call Get_Gen_Memo_Recs(aryGenText(intSub,1))
             strDesc = aryGenMemo(1,3)
             intCntr = intCntr + 1
             aryTxtRecs(intCntr) = "<br>"
             intCntr = intCntr + 1
             If left(strPicFile,5) <> "http:" and left(strPicFile,1) <> "/" Then
                strPicFile = "/" & APPLICATION("PICDIR") & "/" & strPicFile
             End If
             aryTxtRecs(intCntr) = "<img src = '" & strPicFile & "'><br>"
             If strCaption <> "" Then
                intCntr = intCntr + 1
                aryTxtRecs(intCntr) = "<font size=4><b>" & strCaption & "</b></font><br>"
             End If
             If strDesc <> "" Then
                intCntr = intCntr + 1
                aryTxtRecs(intCntr) = strDesc & "<br>"
             End If
          End If
       Next

       aryTxtRecs(0) = intCntr
       call Write_Text_File(strWebName & ".ssi",Application("INCLDIR"))
       If strWebSite <> "" Then
          Response.Redirect strWebSite
       Else
          Response.END
       End If
    End If

    If strWebName <> "" Then
       Call Get_Gen_Text_Recs("PHOTO",ucase(strWebName))
       Call Get_Gen_Text2_Recs(ucase(strWebName))
    Else
       aryGenText(0,0) = 0
       aryGenText2(0,0) = 0
    End If

    If aryGenText(0,0) = 0 Then
       aryGenText(0,0) = aryGenText(0,0) + 5
    Else
       aryGenText(0,0) = aryGenText(0,0) + 2
    End If

    strTblClr = "#b0c4de"
    strFSize = "+1"
    call Setup_Web_Page("Update " & strTitle  & " Photo Details",4)

    Response.Write "<tr><td><b>Seq</b></td><td align='center'><b>Photo File /<br>Caption</b></td><td align='center'><b>Description</b></td></tr>" & vbCrLf

    Response.Write "<form action='" & strPgm & "' method='post'>" & vbCrLf
    Response.Write "<input type='hidden' name='UpdFile' value='Y'>" & vbCrLf
    Response.Write "<input type='hidden' name='WN' value='" & strWebName & "'>" & vbCrLf
    Response.Write "<input type='hidden' name='WS' value='" & strWebSite & "'>" & vbCrLf

    If strMenu <> "" Then
        Response.Write "<br><br><input type=" & CHR(34) & "button" & CHR(34) & " value=" & CHR(34) & "Return to Menu" & CHR(34) & " onClick=" & CHR(34) & "document.location.href = '" & strMenu & "'" & CHR(34) & ">" & vbCrLf
    End If

    intCntr = 0
    For intSub = 1 to aryGenText(0,0)
       intCntr = intCntr + 1
       strPicFile = aryGenText(intSub,5)
       If strPicFile = "" Then
          strCaption = ""
          strDesc = ""
          intSeq = ""
       Else
          intSeq = aryGenText(intSub,4)
          strCaption = GetPhotoCaption(aryGenText(intSub,1),ucase(strWebName))
          Call Get_Gen_Memo_Recs(aryGenText(intSub,1))
          strDesc = aryGenMemo(1,3)
       End If
       Response.Write "<tr><td>" & vbCrLf
       Response.Write "<input type='text' name='Seq" & intCntr & "' size='1' maxlength='2' value=" & CHR(34) & intSeq & CHR(34) & ">" & vbCrLf
       Response.Write "</td><td>" & vbCrLf
       Response.Write "<input type='text' name='Pho" & intCntr & "' size='50' maxlength='100' value=" & CHR(34) & strPicFile & CHR(34) & ">" & vbCrLf
       If strPicFile <> "" Then
          Response.Write "<input type='hidden' name='PID" & intCntr & "' value='" & aryGenText(intSub,1) & "'>" & vbCrLf
       End If
       If strMenu <> "" Then
          Response.Write "<input type='hidden' name='M' value='" & strMenu &"'>" & vbCrLf
       End If
       Response.Write "</td><td rowspan=2>" & vbCrLf
       Response.Write "<textarea name='Dsc" & intCntr & "' cols=60 rows=3 wrap='YES'>" & vbCrLf
       Response.Write strDesc & vbCrLf
       Response.Write "</textarea>" & vbCrLf
       If strDesc <> "" Then
          Response.Write "<input type='hidden' name='DID" & intCntr & "' value='" & aryGenMemo(1,1) & "'>" & vbCrLf
       End If
       Response.Write "</td></tr>" & vbCrLf
       Response.Write "<tr><td>&nbsp;</td><td colspan=2>" & vbCrLf
       Response.Write "<input type='text' name='Cap" & intCntr & "' size='50' maxlength='100' value='" & strCaption & "'>" & vbCrLf
       If strCaption <> "" Then
          Response.Write "<input type='hidden' name='CID" & intCntr & "' value='" & intCID & "'>" & vbCrLf
       End If
       Response.Write "</td></tr>" & vbCrLf
    Next
    SESSION("CNT") = intCntr


    Response.Write "<tr><td colspan='3'>" & vbCrLf
    Response.Write "<hr width='100%' size='5' noshade>" & vbCrLf
    Response.Write "</td></tr>" & vbCrLf
    Response.Write "<tr><td align='center' colspan='3'>" & vbCrLf
    Response.Write "<input type='SUBMIT' name='Update' value='Update'>" & vbCrLf
    Response.Write "</td></tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "<div align='center'>" & vbCrLf
    Response.Write "<font color='red' size=-1>" & vbCrLf
    Response.Write "Picture Files Not Prefixed With a Location Need to be Stored in the folder " & CHR(34) & "/" & APPLICATION("PICDIR") & "/" & CHR(34) & "<br>"
    Response.Write "Leave the Photo File Name Blank to Delete the Entry" & vbCrLf
    Response.Write "</font>" & vbCrLf
    Response.Write "</div></form>" & vbCrLf

FUNCTION GetPhotoCaption(ID,CAT)

    DIM intSub

    For intSub = 1 to aryGenText2(0,0)
       If aryGenText2(intSub,8) = CAT and aryGenText2(intSub,6) = ID Then
         GetPhotoCaption = aryGenText2(intSub,4)
         intCID = aryGenText2(intSub,1)
         EXIT FUNCTION
       End If
    Next

END FUNCTION
%>


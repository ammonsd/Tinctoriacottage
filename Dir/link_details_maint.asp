<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/maint_webpage.inc" -->
<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->
<!--#INCLUDE FILE="../dir/links_maint_setup.inc" -->
<!--#INCLUDE FILE="../dir/db_gen_select.inc" -->
<!--#INCLUDE FILE="../dir/db_gen_maint.inc" -->
<!--#INCLUDE FILE="../dir/link_group_select.inc" -->

<%
    DIM intSub, intSub2, intCntr, strPgm, strHld, X, bolUpOK, intFKey2, strGrpName, strFKey
    DIM strLnkID, intLnkID, intLnkDID, intSeq, strWebName, strLink, strLinkDesc, strSeqHld, strMenu
    DIM aryLinkDsc(500), strServerName

    strPgm = "link_details_maint.asp"
    strServerName=lcase(Request.ServerVariables("SERVER_NAME"))

    If Request.QueryString("FK") <> "" Then
       strFkey = Request.QueryString("FK")
    Else
       strFkey = Request.Form("FK")
    End If
    If instr(strFkey,",") > 0 Then
       strFkey = left(strFkey,instr(strFkey,",")-1)
    End If
    If strFKey <> "" and isNumeric(strFkey) Then
       intFKey = int(strFKey)
    End If
    strGrpName = Request.QueryString("GRP")
    If instr(strGrpName,",") > 0 Then
       strGrpName = left(strGrpName,instr(strGrpName,",")-1)
    End If

    If Request.QueryString("NW") <> "" Then
       SESSION("NW") = Request.QueryString("NW")
    End If

    strWebName = Request.QueryString("WN")
    If strWebName = "" Then
       strWebName = Request.Form("WN")
    End If
    If strWebName = "" Then
       strWebName = "linkspg"
    End If
    strMenu = Request.QueryString("M")
    If strMenu = "" Then
       strMenu = Request.Form("M")
    End If

    Call System_Setup("NONE")
    Call Logon_Check(GetCurPath("")&strPgm,3,strLogonGrp)
    Call Database_Setup

    If Request.QueryString("BLD") = "Y" Then
       Call Build_SSI_File(Request.QueryString("WS"))
    End If

    If Request.Form("UpdFile") = "Y" Then
       If Request.Form("AlphaSort") = "" Then
          aryLinkDsc(0) = 0
       Else
          intCntr = 0
          For intSub = 1 to SESSION("CNT")
             X = "Lnk" & intSub
             strLink = Request.Form(X)
             If strLink <> "" Then
                X = "Dsc" & intSub
                strLinkDesc = Request.Form(X)
                X = "LnkID" & intSub
                aryLinkDsc(intCntr) = strLinkDesc & " " & Request.Form(X)
                intCntr = intCntr + 1
             End If
          Next
          arySorted = SortArray(aryLinkDsc,0)

          intCntr = 0
          For intSub = 0 to intNbrEntries - 1
             intCntr = intCntr + 1
             aryLinkDsc(intCntr) = arySorted(intSub)
             aryLinkDsc(0) = intCntr
          Next
          Erase arySorted
       End If
       intFKey = int(Request.Form("CFK"))
       intLastSeq = 0
       strSeqHld = ":"
       bolDoPhrases = false
       For intSub = 1 to SESSION("CNT")
          X = "Lnk" & intSub
          strLink = Request.Form(X)
          X = "LnkID" & intSub
          strLnkID = Request.Form(X)
          X = "Dsc" & intSub
          strLinkDesc = Request.Form(X)
          X = "LnkDID" & intSub
          If Request.Form(X)  = "" or not isNumeric(Request.Form(X)) Then
             intLnkDID = 0
          Else
             intLnkDID = int(Request.Form(X))
          End If
          If aryLinkDsc(0) > 0 Then
             intSeq = intLastSeq + 1
             For intSub2 = 1 to aryLinkDsc(0)
                If GetWords(aryLinkDsc(intSub2),TotalWords(aryLinkDsc(intSub2)),1) = strLnkID Then
                   intSeq = intSub2
                   EXIT FOR
                End If
             Next
          Else
             X = "Seq" & intSub
             If Request.Form(X) <> "" and instr(strSeqHld,":"&Request.Form(X)&":") = 0 Then
                intSeq = int(Request.Form(X))
             Else
                intSeq = intLastSeq + 1
             End If
             If Request.Form("AlphaSort") <> "" Then
                intSeq = 1
             End If
          End If
          If intSeq > intLastSeq Then
             intLastSeq = intSeq
          End If
          strSeqHld = strSeqHld & intSeq & ":"
          If strLink <> "" or strLnkID <> "" Then
             If strLnkID = "" or not isNumeric(strLnkID) Then
                intLnkID = 0
             Else
                intLnkID = int(strLnkID)
             End If
             If intLnkID = 0 Then
                intFKey2 = AddGenTextRec("LINK",ucase(strWebName),intSeq,strLink,false,intFKey,"")
                If strLinkDesc <> "" Then
                   Call AddGenTextRec("LINKDESC",ucase(strWebName),intSeq,strLinkDesc,false,intFKey,intFKey2)
                End If
             ElseIf intLnkID > 0 and (strLinkDesc <> "" and intLnkDID = 0) Then
                   Call AddGenTextRec("LINKDESC",ucase(strWebName),intSeq,strLinkDesc,false,intFKey,intLnkID)
             ElseIf strLink = "" Then
                strLink = ""
                Call CascadeDelGenTextRec(intLnkID)
             ElseIf strLinkDesc = "" and intLnkDID > 0 Then
                Call DelGenTextRec(intLnkDID)
                intLnkDID = 0
             End If
             If intLnkID > 0 and strLink <> "" Then
                Call UpdGenTextRec(intLnkID,ucase(strWebName),intSeq,strLink,false)
                If intLnkDID > 0 Then
                   Call UpdGenTextRec(intLnkDID,ucase(strWebName),intSeq,strLinkDesc,false)
                End If
             End If
          End If
       Next

       Response.Write "<div align='center'>" & vbCrLf
       Response.Write "<br><br><b><br><br><b><font color='red'>Link Details Updated</font></b>" & vbCrLf
       If strMenu <> "" Then
           Response.Write "<br><br><input type=" & CHR(34) & "button" & CHR(34) & " value=" & CHR(34) & "Return to Menu" & CHR(34) & " onClick=" & CHR(34) & "document.location.href = '" & strMenu & "'" & CHR(34) & ">" & vbCrLf
       End If
       Response.Write "</div'>" & vbCrLf
       SESSION("CNT") = ""
       Response.End
    End If

    strTblClr = "#b0c4de"
    strFSize = "+1"

    If strGrpName <> "" Then
       strHld = "Update " & strGrpName & " Links"
    Else
       strHld = "Update Link Details"
    End If
    call Setup_Web_Page(strHld,3)

    Response.Write "<tr><td colspan=3>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>Select Link Group</b><br>" & vbCrLf
    Call Build_Link_Group_Select(intFKey,strPgm&"?FK=")
    Response.Write "</td></tr>" & vbCrLf

    If intFkey <> "" and isNumeric(intFKey) Then
       Call Get_Gen_Text2_Recs(intFKey)
    Else
       aryGenText2(0,0) = 0
    End If
    aryGenText2(0,0) = aryGenText2(0,0) + 5

    Response.Write "<tr><td><b>Seq</b></td><td align='center'><b>Description</b></td><td><b>Link</b></td></tr>" & vbCrLf

    Response.Write "<form action='" & strPgm & "' method='post'>" & vbCrLf
    Response.Write "<input type='hidden' name='UpdFile' value='Y'>" & vbCrLf
    Response.Write "<input type='hidden' name='CFK' value='" & intFKey & "'>" & vbCrLf
    Response.Write "<input type='hidden' name='WN' value='" & strWebName & "'>" & vbCrLf

    If strMenu <> "" Then
       Response.Write "<input type='hidden' name='M' value='" & strMenu &"'>" & vbCrLf
    End If

    intCntr = 0
    For intSub = 1 to aryGenText2(0,0)
       If aryGenText2(intSub,2) = "LINK" or aryGenText2(intSub,2) = "" Then
          intCntr = intCntr + 1
          strLink = aryGenText2(intSub,4)
          Response.Write "<tr><td>" & vbCrLf
          Response.Write "<input type='text' name='Seq" & intCntr & "' size='2' maxlength='2' value='" & aryGenText2(intSub,3) & "'>" & vbCrLf
          Response.Write "</td><td>" & vbCrLf
          strLinkDesc = GetLinkDesc(aryGenText2(intSub,1))
          Response.Write "<input type='text' name='Dsc" & intCntr & "' size='70' maxlength='100' value=" & CHR(34) & strLinkDesc & CHR(34) & ">" & vbCrLf
          If strLinkDesc <> "" Then
             Response.Write "<input type='hidden' name='LnkDID" & intCntr & "' value='" & intLnkDID & "'>" & vbCrLf
          End If
          Response.Write "</td><td>" & vbCrLf
          Response.Write "<input type='text' name='Lnk" & intCntr & "' size='70' maxlength='100' value=" & CHR(34) & strLink & CHR(34) & ">" & vbCrLf
          If strLink <> "" Then
             Response.Write "<input type='hidden' name='LnkID" & intCntr & "' value='" & aryGenText2(intSub,1) & "'>" & vbCrLf
          End If
          Response.Write "</td></tr>" & vbCrLf
       End If
    Next
    SESSION("CNT") = intCntr

    Response.Write "<tr><td colspan='3'>" & vbCrLf
    Response.Write "<hr width='100%' size='5' noshade>" & vbCrLf
    Response.Write "</td></tr>" & vbCrLf
    Response.Write "<tr><td align='center' colspan='3'>" & vbCrLf
    Response.Write "<input type='SUBMIT' name='Update' value='Update Sequentially'>" & vbCrLf
    Response.Write "<input type='SUBMIT' name='AlphaSort' value='Update Alphabetically'>" & vbCrLf
    Response.Write "</td></tr>" & vbCrLf
    Response.Write "</table>"
    Response.Write "<div align='center'>" & vbCrLf
    Response.Write "<font color='red' size=-1>Leave Link Blank to Remove Entry</font>" & vbCrLf
    Response.Write "</div></form>" & vbCrLf

FUNCTION GetLinkDesc(ID)

    DIM intSub

    For intSub = 1 to aryGenText2(0,0)
       If aryGenText2(intSub,2) = "LINKDESC" and aryGenText2(intSub,7) = ID Then
         GetLinkDesc = aryGenText2(intSub,4)
         intLnkDID = aryGenText2(intSub,1)
         EXIT FUNCTION
       End If
    Next

END FUNCTION

SUB Build_SSI_File(WEBSITE)

    DIM intSub, intSub2, intCntr, strSSIFn

    strSSIFn = strWebName & ".ssi"

    bolNoSrch = true

    Call Get_Gen_Text_Recs("LINKGRP","LINKSPG")

    intCntr = 1
    aryTxtRecs(intCntr) = "<br>"

    For intSub = 1 to aryGenText(0,0)
       intCntr = intCntr + 1
       aryTxtRecs(intCntr) = "<br>"
       intCntr = intCntr + 1
       aryTxtRecs(intCntr) = "<img src='/" & Application("PGMDIR") & "/graphics/spacer.gIf' width='10' height='1' border='0' alt=''>"
       intCntr = intCntr + 1
       aryTxtRecs(intCntr) = "&nbsp;&nbsp;"
       intCntr = intCntr + 1
       aryTxtRecs(intCntr) = "<b>" & aryGenText(intSub,5) & "</b><br>"

       Call Get_Gen_Text2_Recs(aryGenText(intSub,1))
       For intSub2 = 1 to aryGenText2(0,0)
          If aryGenText2(intSub2,2) = "LINK" Then
             If ucase(left(aryGenText2(intSub2,4),5)) <> "HTTP:" Then
                aryGenText2(intSub2,4) = "http://" & aryGenText2(intSub2,4)
             End If
             intCntr = intCntr + 1
             aryTxtRecs(intCntr) = "<img src='/" & Application("PGMDIR") & "/graphics/spacer.gIf' width='48' height='1' border='0' alt=''>"
             intCntr = intCntr + 1
             aryTxtRecs(intCntr) = "<a href=" & aryGenText2(intSub2,4)
             If SESSION("NW") = "Y" and instr(lcase(aryGenText2(intSub2,4)),strServerName) = 0 Then
                aryTxtRecs(intCntr) = aryTxtRecs(intCntr) & " target='_blank'"
             End If
             strLinkDesc = GetLinkDesc(aryGenText2(intSub2,1))
             If strLinkDesc = "" Then
                strLinkDesc = aryGenText2(intSub2,4)
             End If
             aryTxtRecs(intCntr) = aryTxtRecs(intCntr) & " onmouseover=" & CHR(34) & "window.status='" & strLinkDesc & "';return true" & CHR(34) & ">" & strLinkDesc & "</a><br>"
          End If
       Next
    Next

    aryTxtRecs(0) = intCntr
    call Write_Text_File(strSSIFn,Application("INCLDIR"))
    If WEBSITE <> "" Then
       Response.Redirect WEBSITE
    End If
    Response.End

END SUB
%>


<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/maint_webpage.inc" -->
<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/db_setup.inc" -->
<!--#INCLUDE FILE="../dir/process_setup.inc" -->
<!--#INCLUDE FILE="../dir/db_gen_select.inc" -->
<!--#INCLUDE FILE="../dir/db_gen_maint.inc" -->

<%
    DIM intSub, intCntr, strPgm, strHld, strIncFn, X, intSub2, aryInc(100,2), bolUpOK
    DIM strLinkGrp, strID, intID, intSeq, strMenu

    strPgm = "link_group_maint.asp"

    Call System_Setup("NONE")
    Call Logon_Check(GetCurPath("")&strPgm,3,strLogonGrp)

    Call Database_Setup
    strMenu = Request.QueryString("M")
    If strMenu = "" Then
       strMenu = Request.Form("M")
    End If

    strIncFn = "setup_link_groups.inc"

    If Request.Form("UpdFile") = "Y" Then
       intLastSeq = 0
       intSub2 = 0
       For intSub = 1 to SESSION("CNT")
          X = "Grp" & intSub
          strLinkGrp = Request.Form(X)
          X = "ID" & intSub
          strID = Request.Form(X)
          X = "Seq" & intSub
          If Request.Form(X) <> "" Then
             intSeq = int(Request.Form(X))
          Else
             intSeq = intLastSeq + 1
          End If
          If intSeq > intLastSeq Then
             intLastSeq = intSeq
          End If
          If strLinkGrp <> "" or strID <> "" Then
             If strID = "" or not isNumeric(strID) Then
                intID = 0
             Else
                intID = int(strID)
             End If
             If intID = 0 Then
                intSub2 = intSub2 + 1
                intID = AddGenTextRec("LINKGRP","LINKSPG",intSeq,strLinkGrp,false,"","")
             ElseIf strLinkGrp = "" Then
                strLinkGrp = ""
                Call CascadeDelGenTextRec(intID)
             Else
                intSub2 = intSub2 + 1
                Call UpdGenTextRec(intID,"LINKSPG",intSeq,strLinkGrp,false)
             End If
             If strLinkGrp <> "" Then
                aryInc(intSub2,1) = strLinkGrp
                aryInc(intSub2,2) = intID
             End If
          End If
       Next
       aryInc(0,0) = intSub2

       intSub = 1
       aryTxtRecs(intSub) = "<%SUB Build_Link_Group_Array"
       For intSub2 = 1 to aryInc(0,0)
          intSub = intSub + 1
          aryTxtRecs(intSub) = "aryLinkGroup(" & intSub2 & ",1) = " & CHR(34) &  aryInc(intSub2,1) & CHR(34)
          intSub = intSub + 1
          aryTxtRecs(intSub) = "aryLinkGroup(" & intSub2 & ",2) = " &  aryInc(intSub2,2)
       Next
       intSub = intSub + 1
       aryTxtRecs(intSub) = "aryLinkGroup(0,0) = " & aryInc(0,0)
       intSub = intSub + 1
       aryTxtRecs(intSub) = "END SUB" & CHR(37) & ">"
       aryTxtRecs(0) = intSub

       call Write_Text_File(strIncFn,Application("INCLDIR"))

       Response.Write "<div align='center'>" & vbCrLf
       Response.Write "<br><br><b><br><br><b><font color='red'>Link Groups Updated</font></b>" & vbCrLf
       If strMenu <> "" Then
           Response.Write "<br><br><input type=" & CHR(34) & "button" & CHR(34) & " value=" & CHR(34) & "Return to Menu" & CHR(34) & " onClick=" & CHR(34) & "document.location.href = '" & strMenu & "'" & CHR(34) & ">" & vbCrLf
       End If
       Response.Write "</div'>" & vbCrLf
       SESSION("CNT") = ""
       Response.End
    End If

    Call Get_Gen_Text_Recs("LINKGRP","LINKSPG")

    SESSION("CNT") = aryGenText(0,0) + 5

    strTblClr = "#b0c4de"
    strFSize = "+1"
    call Setup_Web_Page("Update Link Groups",2)

    Response.Write "<tr><td><b>Seq</b></td><td align='center'><b>Link Group</b></td></tr>" & vbCrLf

    Response.Write "<form action='link_group_maint.asp' method='post'>" & vbCrLf
    Response.Write "<input type='hidden' name='UpdFile' value='Y'>" & vbCrLf
    If strMenu <> "" Then
       Response.Write "<input type='hidden' name='M' value='" & strMenu &"'>" & vbCrLf
    End If

    For intSub = 1 to SESSION("CNT")
       Response.Write "<tr><td>" & vbCrLf
       Response.Write "<input type='text' name='Seq" & intSub & "' size='2' maxlength='2' value='" & aryGenText(intSub,4) & "'>" & vbCrLf
       Response.Write "</td><td>" & vbCrLf
       Response.Write "<input type='text' name='Grp" & intSub & "' size='80' maxlength='80' value='" & aryGenText(intSub,5) & "'>" & vbCrLf
       Response.Write "<input type='hidden' name='ID" & intSub & "' value='" & aryGenText(intSub,1) & "'>" & vbCrLf
       Response.Write "</td></tr>" & vbCrLf
    Next

    Response.Write "<tr><td colspan='2'>" & vbCrLf
    Response.Write "<hr width='100%' size='5' noshade>" & vbCrLf
    Response.Write "</td></tr>" & vbCrLf
    Response.Write "<tr><td align='center' colspan='2'>" & vbCrLf
    Response.Write "<input type='SUBMIT' name='Update' value='Update'>" & vbCrLf
    Response.Write "</td></tr>" & vbCrLf
    Response.Write "</table>"
    Response.Write "<div align='center'>" & vbCrLf
    Response.Write "<font color='red' size=-1>Leave Group Name Blank to Remove Entry and <b>ALL</b> Associated Links</font>" & vbCrLf
    Response.Write "</div></form>" & vbCrLf
%>


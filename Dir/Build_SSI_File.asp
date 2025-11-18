<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>
<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<%

    DIM aryRecs(1000), intSub, strPgm, bolDscFnd, strReqLogonDir, strSSIFn, strGoTo, aryErrs(20), intErr

    strPgm = "build_ssi_file.asp"

    If SESSION("BPG") = "" Then
       SESSION("BPG") = 1
    End If

    Call System_Setup("NONE")

    If strLogonGrp <> "" Then
       strReqLogonDir = strLogonGrp
    Else
       strReqLogonDir = "TXTFN"
    End If

    Call Logon_Check(GetCurPath("")&strPgm,3,strReqLogonDir)

    intErr = 0
    aryErrs(0) = intErr

    If SESSION("PGM") <> "" and Request.QueryString("FN") = "" Then
       call Read_Text_File(SESSION("FN"),Application("FILEDIR"))
       strSSIFn = left(SESSION("FN"),instr(SESSION("FN"),".")) & "ssi"
       If SESSION("TY") = "L" Then
          call Build_Link_Content()
       ElseIf SESSION("TY") = "B" Then
          call Build_Book_Content()
       Else
          call Build_Photo_Content()
       End If
       If aryErrs(0) > 0 Then
          Response.Write "The Following Photo Files Were Not Found in <b>" & GetPgmPath(APPLICATION("PICDIR")) & "</b><br><br>"
          For intErr = 1 to aryErrs(0)
             Response.Write aryErrs(intErr) & "<br>"
          Next
          Response.End
       End If
       For intSub = 1 to aryRecs(0)
          aryTxtRecs(intSub) = aryRecs(intSub)
       Next
       aryTxtRecs(0) = intSub
       call Write_Text_File(strSSIFn,Application("INCLDIR"))
       strGoTo = SESSION("WP")
       SESSION("FN") = ""
       SESSION("TI") = ""
       SESSION("PGM") = ""
       SESSION("TY") = ""
       SESSION("WP") = ""
       SESSION("HF") = ""
       SESSION("NW") = ""
       If strGoTo <> "" Then
          Response.Redirect strGoTo
       End If
       Response.Write "<br><br><dev align='center'><b>File " & strSSIFn & " Updated</b></div>"
    Else
       SESSION("FN") = Request.QueryString("FN")
       SESSION("TY") = Request.QueryString("TY")
       SESSION("NW") = Request.QueryString("NW")
       SESSION("TI") = ""
       SESSION("WP") = ""
       SESSION("PGM") = strPgm
       strSSIFn = ucase(left(SESSION("FN"),instr(SESSION("FN"),".")-1))
       call Read_Text_File("Build_SSI.dsc",Application("FILEDIR"))
       For intSub = 1 to aryTxtRecs(0)
          If left(ucase(aryTxtRecs(intSub)),2)="H:" Then
             SESSION("HF") = mid(aryTxtRecs(intSub),3)
          ElseIf left(aryTxtRecs(intSub),1) <> " " and left(ucase(aryTxtRecs(intSub)),2)<>"W:" Then
             If ucase(aryTxtRecs(intSub)) = strSSIFn Then
                bolDscFnd = true
             Else
                bolDscFnd = false
             End If
          ElseIf bolDscFnd Then
             If left(ucase(aryTxtRecs(intSub)),2)="W:" Then
                SESSION("WP") = mid(aryTxtRecs(intSub),3)
             ElseIf SESSION("TI") = "" Then
                SESSION("TI") = mid(aryTxtRecs(intSub),2)
             Else
                If left(trim(aryTxtRecs(intSub)),1) <> "<" Then
                   SESSION("TI") = SESSION("TI") & "<br>"
                End If
                SESSION("TI") = SESSION("TI") & mid(aryTxtRecs(intSub),2)
             End If
          End If
       Next
       Response.Redirect "maint_text_file.asp?E=" & Request.QueryString("E")
    End If

SUB Build_Photo_Content()

    DIM intSub, intSub2, bolHdr

    intSub = 1
    aryRecs(intSub) = "<br>"
    For intSub2 = 1 to aryTxtRecs(0)
       If right(ucase(aryTxtRecs(intSub2)),4) = ".GIF" or right(ucase(aryTxtRecs(intSub2)),4) = ".JPG" Then
          intSub = intSub + 1
          aryRecs(intSub) = "<br>"
          intSub = intSub + 1
          aryRecs(intSub) = "<img src = '/" & APPLICATION("PICDIR") & "/" & aryTxtRecs(intSub2) & "'><br>"
          bolHdr = true
          If CheckFileExist("/" & APPLICATION("PICDIR") & "/" & aryTxtRecs(intSub2)) <> 0 Then
             intErr = intErr + 1
             aryErrs(intErr) = aryTxtRecs(intSub2)
             aryErrs(0) = intErr
          End If
       ElseIf bolHdr Then
          If trim(aryTxtRecs(intSub2)) <> "" Then
             intSub = intSub + 1
             aryRecs(intSub) = "<font size=4><b>" & aryTxtRecs(intSub2) & "</b></font><br>"
          End If
          bolHdr = false
       ElseIf trim(aryTxtRecs(intSub2)) <> "" Then
          intSub = intSub + 1
          aryRecs(intSub) = aryTxtRecs(intSub2) & "<br>"
       End If
    Next
    aryRecs(0) = intSub

END SUB

SUB Build_Link_Content()

    DIM intSub, intSub2, strLink, strDesc, intLoc

    intSub = 1
    aryRecs(intSub) = "<br>"
    For intSub2 = 1 to aryTxtRecs(0)
       If trim(aryTxtRecs(intSub2)) <> "" Then
          If left(ucase(aryTxtRecs(intSub2)),5) = "HTTP:" Then
             intLoc = instr(aryTxtRecs(intSub2)," ")
             If intLoc = 0 Then
                strLink = aryTxtRecs(intSub2)
                strDesc = strLink
             Else
               strLink = left(aryTxtRecs(intSub2),intLoc-1)
               strDesc = mid(aryTxtRecs(intSub2),intLoc+1)
             End If
             intSub = intSub + 1
             aryRecs(intSub) = "<img src='/" & Application("PGMDIR") & "/graphics/spacer.gIf' width='48' height='1' border='0' alt=''>"
             intSub = intSub + 1
             aryRecs(intSub) = "<a href=" & strLink
             If SESSION("NW") = "Y" Then
                aryRecs(intSub) = aryRecs(intSub) & " target='_blank'"
             End If
             aryRecs(intSub) = aryRecs(intSub) & ">" & strDesc & "</a><br>"
          Else
             intSub = intSub + 1
             aryRecs(intSub) = "<br>"
             intSub = intSub + 1
             aryRecs(intSub) = "<img src='/" & Application("PGMDIR") & "/graphics/spacer.gIf' width='10' height='1' border='0' alt=''>"
'             intSub = intSub + 1
'             aryRecs(intSub) = "<img src='/" & Application("PGMDIR") & "/graphics/Red_Dot.gIf' width='12' height='12' border='0' alt=''>"
             intSub = intSub + 1
             aryRecs(intSub) = "&nbsp;&nbsp;"
             intSub = intSub + 1
             aryRecs(intSub) = "<b>" & aryTxtRecs(intSub2) & "</b><br>"
          End If
       End If
    Next
    aryRecs(0) = intSub

END SUB

SUB Build_Book_Content()

    DIM intSub, intSub2, bolHdr, strPhoto

    intSub = 1
    aryRecs(intSub) = "<table>"
    For intSub2 = 1 to aryTxtRecs(0)
       intSub = intSub + 1
       aryRecs(intSub) = "<tr>"
       If right(ucase(aryTxtRecs(intSub2)),4) = ".GIF" or right(ucase(aryTxtRecs(intSub2)),4) = ".JPG" Then
          strPhoto = "<td><img src = '/" & APPLICATION("PICDIR") & "/" & aryTxtRecs(intSub2) & "'></td>"
          bolHdr = true
          If CheckFileExist("/" & APPLICATION("PICDIR") & "/" & aryTxtRecs(intSub2)) <> 0 Then
             intErr = intErr + 1
             aryErrs(intErr) = aryTxtRecs(intSub2)
             aryErrs(0) = intErr
          End If
          intSub2 = intSub2 + 1
       Else
          strPhoto = "<td>&nbsp;</td>"
       End If
       intSub = intSub + 1
       aryRecs(intSub) = "<td><font size=4><b>" & aryTxtRecs(intSub2) & "</b></font><br>"
       intSub2 = intSub2 + 1
       intSub = intSub + 1
       aryRecs(intSub) = "By " & aryTxtRecs(intSub2) & "<br><br>"
       intSub2 = intSub2 + 1
       intSub = intSub + 1
       aryRecs(intSub) = aryTxtRecs(intSub2) & "</td>"
       intSub = intSub + 1
       aryRecs(intSub) = strPhoto & "</tr>"
       intSub = intSub + 1
       aryRecs(intSub) = "<tr><td>&nbsp;</td></tr>"
    Next
    intSub = intSub + 1
    aryRecs(intSub) = "</table>"
    aryRecs(0) = intSub

END SUB

%>


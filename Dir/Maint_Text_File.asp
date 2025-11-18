<%@ LANGUAGE="VBSCRIPT" %>

<%
OPTION EXPLICIT
Response.Buffer = True
%>

<!--#INCLUDE FILE="../dir/maint_webpage.inc" -->
<!--#INCLUDE FILE="../dir/write_errors.inc" -->
<!--#INCLUDE FILE="../dir/form_data.inc" -->
<!--#INCLUDE FILE="../dir/logon_check.inc" -->
<!--#INCLUDE FILE="../dir/get_data.inc" -->
<!--#INCLUDE FILE="../dir/cookie_maint.inc" -->
<!--#INCLUDE FILE="../dir/get_folder_details.inc" -->
<!--#INCLUDE FILE="../dir/process_setup.inc" -->

<%
    DIM intSub, strReqFn, strReqTitle, intCntr, strPgm, strHld, strReqLogonDir, intSL, intRows, intBPG

    strPgm = "maint_text_file.asp"

    Call System_Setup("NONE")

    If SESSION("RQL") <> "" Then
       intSL = SESSION("RQL")
    Else
       intSL = 5
    End If

    If strLogonGrp <> "" Then
       strReqLogonDir = strLogonGrp
    Else
       strReqLogonDir = "TXTFN"
    End If

    Call Logon_Check(GetCurPath("")&strPgm,intSL,strReqLogonDir)

    strGetData = SESSION("GETDATA")
    Session.Contents.Remove("GETDATA")

    If SESSION("BPG") <> "" Then
       intBPG = int(SESSION("BPG"))
       SESSION("BPG") = ""
    Else
       intBPG = 1
    End If

    strReqFn=trim(FormData("FN"))
    strReqTitle=FormData("TI")

    intCntr = instr(strReqFn,"[CD]")
    If intCntr > 0 Then
       strHld = SESSION("HOMEDIR")
       If strHld = "" Then
          strHld = strCurDir
       End If
       strReqFn = left(strReqFn,intCntr-1) & strHld & mid(strReqFn,intCntr+4)
    End If

    If strReqFn = "" and strGetData <> "Y" Then
       aryData(1,1) = "Name of Text File"
       aryData(2,1) = Request.Cookies("TxtFn_"&SESSION("USERID"))("LastFile")
       aryData(3,1) = ""
       aryData(4,1) = "FN"
       aryData(0,0) = 1
       call Get_Data(strPgm,"Text File Name",30)
       Response.End
    ElseIf strGetData = "Y" Then
       If strReqFn = "" Then
          Response.End
       ElseIf left(strReqFn,1) = "*" Then
          intCntr = 0
          call Get_Folder_File_Details("cgi-bin",strReqFn)
          For intSub = 1 to aryFDtls(0,0)
            If trim(aryFDtls(1,intSub)) <> "" Then
               intCntr = intCntr + 1
               aryData(1,intSub) = aryFDtls(1,intSub)
               aryData(2,intSub) = aryFDtls(1,intSub)
               aryData(3,intSub) = "R"
               aryData(4,intSub) = "FN"
            End If
          next
          aryData(0,0) = intCntr
          strFldAlign = "R"
          call Get_Data(strPgm,"Select File to Process",0)
          Response.End
       Else
          Call AddCookie("TxtFn_"&SESSION("USERID"), "LastFile", strReqFn, 365)
       End If
    End If

    bolNoTable = true
    bolNoLinks = true

    If strReqTitle = "" Then
       strReqTitle = "<div align='center'>Update File " & strReqFn & "</div>"
    End If
    intRows = 32 - (CountChars(strReqTitle,"<br>")+1)
    intRows = intRows - (CountChars(strReqTitle,"<li>")+1)
    call Setup_Web_Page(strReqTitle,0)

    If Request.Form("Update") <> "" Then
       strHld = trim(Request.Form("Text"))
       For intCntr = len(strHld) to 1 step -1
          If ASC(mid(strHld,intCntr,1)) > 33 Then
             strHld = left(strHld,intCntr)
             EXIT FOR
          End If
       Next
       aryTxtRecs(1) = strHld
       aryTxtRecs(0) = 1
       call Write_Text_File(strReqFn,"CGI-BIN")
       If SESSION("PGM") <> "" Then
          Response.Redirect SESSION("PGM")
       End If
    End If

    If bolDelFile Then
       Response.Write "<div align='center'>" & vbCrLf
       Response.Write "<br><br><b><br><br><b><font color='red'>" & CHR(34) & strReqFn & CHR(34) & " Has Been Deleted</font></b>" & vbCrLf
    ElseIf Request.Form("Update") <> "" Then
       Response.Write "<div align='center'>" & vbCrLf
       Response.Write "<br><br><b><br><br><b><font color='red'>" & CHR(34) & strReqFn & CHR(34) & " Was Updated</font></b>" & vbCrLf
    Else
       Response.Write "<div align='center'>" & vbCrLf
       Response.Write "<table><tr><td>" & vbCrLf
       Response.Write "<b>" & strReqTitle & "</b>" & vbCrLf
       Response.Write "</td></tr></table>" & vbCrLf
       If Request.QueryString("E") <> "Y" Then
          call Read_Text_File(strReqFn,"CGI-BIN")
       ElseIf Request.QueryString("MSG") <> "" Then
          aryTxtRecs(0) = 1
          aryTxtRecs(1) = Request.QueryString("MSG")
       End If
       Response.Write "<form action='Maint_Text_File.asp' method='post'>" & vbCrLf
       If SESSION("PGM") = "" Then
          Response.Write "<input type='hidden' name='FN' value='" & strReqFn &"'>" & vbCrLf
          If FormData("TI") <> "" Then
             Response.Write "<input type='hidden' name='TI' value='" & FormData("TI") &"'>" & vbCrLf
          End If
       End If
       Response.Write "<textarea cols='110' rows='" & intRows & "' name='Text'>" & vbCrLf
       For intSub = 1 to aryTxtRecs(0)
          Response.Write aryTxtRecs(intSub) & vbCrLf
       Next
       Response.Write "</textarea><br>" & vbCrLf
       Response.Write "<input type='SUBMIT' name='Update' value='Update'>" & vbCrLf
       If SESSION("HF") <> "" Then
         Response.Write "<img src='/" & Application("PGMDIR") & "/graphics/spacer.gif' width='50' height='0' border='0' alt=''>" & vbCrLf
         Response.Write "<input type=" & CHR(34) & "button" & CHR(34) & " value=" & CHR(34) & "Help" & CHR(34) & " onClick=" & CHR(34) & "window.open('" & SESSION("HF") & "')" & CHR(34) & ">" & vbCrLf
       End If
       Response.Write "<img src='/" & Application("PGMDIR") & "/graphics/spacer.gif' width='50' height='0' border='0' alt=''>" & vbCrLf
       Response.Write "<input type='RESET' name='Reset'>" & vbCrLf
       Response.Write "<img src='/" & Application("PGMDIR") & "/graphics/spacer.gif' width='50' height='0' border='0' alt=''>" & vbCrLf
       Response.Write "<input type='button' value='Cancel' onClick='history.go(-" & intBPG & ")'>"
       Response.Write "</form>" & vbCrLf
    End if
    Response.Write "</div>" & vbCrLf

    call Wrapup_Web_Page
%>



<!--#INCLUDE FILE="../dir/get_program_path.inc" -->
<!--#INCLUDE FILE="../dir/process_text_files.inc" -->
<!--#INCLUDE FILE="../dir/get_words.inc" -->
<!--#INCLUDE FILE="../dir/replace_chars.inc" -->
<!--#INCLUDE FILE="../dir/Local_Setup.inc" -->

<%
    DIM strBrowser, strBrowserVersion, bolCookies, strSpecOpt, bolPerformHC, strCurDir, strCurDirName
    DIM strLogonGrp, strCurSecGrp, strLstKwFn, intActivityCheck, strLogonLoc, strWhrSql
    DIM strSMTPServer, strSupportMail, strSupportName, strLogFn, strFileLoc
    DIM strDirHelpFn, strSrchHelpFn, strMbrHelpFn, strBookHelpFn, strSecurHelpFn

SUB System_Setup(PGM)

    DIM bc, intLoc, strHld

    Call Local_Setup(PGM)
    Call Process_Setup()

    If Application("EnableSrch") = "Y" Then
       strSrchHelpFn = "Search_Help.asp"
    End If
    If not bolSecurityOnly and Application("DIRTBL") <> "" Then
       strDirHelpFn = "/" & Application("PGMDIR") & "/Directory_Documentation.asp"
    End If
    If Application("MBRTBL") <> "" Then
       strMbrHelpFn = "/" & Application("PGMDIR") & "/Member_Documentation.asp"
    End If
    If Application("BOOKTBL") <> "" Then
       strBookHelpFn = "/" & Application("PGMDIR") & "/Book_Documentation.asp"
    End If
    strSecurHelpFn = "/" & Application("PGMDIR") & "/Security_Documentation.asp"

    strFileLoc = APPLICATION("FILELOC")
    If strFileLoc = "" Then
       strFileLoc = "CGI-BIN"
    End If

    If strLogonGrp = "" Then
       strLogonGrp = "DIR"
    End If

    strLogFn = "addr.log"

    strSMTPServer = APPLICATION("STMP")
    If strSMTPServer = "" Then
       strSMTPServer = "mail" & mid(Request.ServerVariables("SERVER_NAME"),instr(Request.ServerVariables("SERVER_NAME"),"."))
    End If

    strSupportName = APPLICATION("SUPPORT_NAME")
    If strSupportName = "" Then
       strSupportName = "Directory Support"
    End if
    strSupportMail = APPLICATION("SUPPORT_MAIL")
    If strSupportMail = "" Then
       strSupportMail = "directory_support@" & mid(Request.ServerVariables("SERVER_NAME"),instr(Request.ServerVariables("SERVER_NAME"),".")+1)
    End if

    strLstKwFn = "list_keywords.asp?TN=" 'Activate List Keywords option on Get Keywords window.
    strLogonLoc = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & Application("PGMDIR") & "/"

    SET bc = Server.CreateObject("MSWC.BrowserType")
    strBrowser = bc.browser
    strBrowserVersion = bc.Version
    bolCookies = bc.cookies

    If strBrowser = "Default" Then
        If instr(Request.ServerVariables("HTTP_USER_AGENT"),"Netscape") > 0 Then
           strBrowser = "Netscape"
        ElseIf instr(Request.ServerVariables("HTTP_USER_AGENT"),"MSIE") > 0 Then
           strBrowser = "IE"
        Else
           strBrowser = "Unknown"
        End If
    End If

    strCurDir = mid(Request.ServerVariables("PATH_INFO"),2)
    intLoc = instr(strCurDir,"/")
    If intLoc > 0 Then
       strCurDir = left(strCurDir,intLoc-1)
    End If
    strCurDirName = ReplChars(strCurDir,"/","")

    strCurSecGrp = GetCurSecGrp()

    If strCurSecGrp <> SESSION("CSG") Then
       SESSION("CSG") = strCurSecGrp
       SESSION("NEWLOGON") = "Y"
    End If

    If Request.QueryString("DEMO") <> ""  Then
       SESSION("DEMO") = Request.QueryString("DEMO")
    End If

    If Request.QueryString("DEBUG") <> ""  Then
       SESSION("DEBUG") = Request.QueryString("DEBUG")
    End If

    If SESSION("HOMEDIR") = "" Then
       SESSION("HOMEDIR") = mid(Request.ServerVariables("SCRIPT_NAME"),2)
       SESSION("HOMEDIR") = left(SESSION("HOMEDIR"),instr(SESSION("HOMEDIR"),"/")-1)
    End If

    If SESSION("MENUFN") = "" or SESSION("NEWLOGON") = "Y" Then
       If SESSION("PTYP") <> "DIR" Then
           SESSION("MENUFN") = "/maint/"
       Else
         SESSION("MENUFN") = Application("MENUFN")
         If SESSION("MENUFN") = "" Then
           If SESSION("GROUP") <> "" Then
              strHld = SESSION("GROUP")
           ElseIf SESSION("SUBLEVL") <> "" Then
              strHld = SESSION("SUBLEVL")
           Else
              strHld = SESSION("HOMEDIR")
           End If
           If strHld <> "" Then
              strHld = strHld & "_"
           End If
           If SESSION("SUBLEVL") <> "" Then
              strHld = SESSION("SUBLEVL") & "/" & strHld
           End If
           SESSION("MENUFN") = "/" & Application("PGMDIR") & "/" & ReplChars(strHld," ","_") & "Menu.htm"
         End If
       End If
    End If

    If Application("MAX_LOGON_INACTIVITY") = "" or not isNumeric(Application("MAX_LOGON_INACTIVITY")) Then
       intActivityCheck = 30
    Else
       intActivityCheck = int(Application("MAX_LOGON_INACTIVITY"))
    End If

    If SESSION("NEWLOGON") = "Y" Then
       Session.Contents.Remove("MENUFN")
       Session.Contents.Remove("KYWRDS")
    End If

    If lcase(PGM) = "none" or PGM = "" Then
       EXIT SUB
    End If

    If Request.Form("BLDLIST") <> "" or bolPerformHC or SESSION("NEWLOGON") = "Y" Then
       Session.Contents.Remove("WHRSQL")
       If SESSION("EMPGM") = "" or bolPerformHC or SESSION("NEWLOGON") = "Y" Then
          Session.Contents.Remove("TO")
          Session.Contents.Remove("CC")
          Session.Contents.Remove("BCC")
          Session.Contents.Remove("EMN")
       End If
    End If

    If lcase(PGM) <> "dir_address_list.asp" and strSpecOpt = "" Then
       Session.Contents.Remove("emKW")
       Session.Contents.Remove("SO")
    End If

END SUB

FUNCTION GetCurSecGrp()

    DIM strSecGrps, strKwGrps, strHld, strKwGrp, strSecGrp, intSub, intSub2

    strSecGrps = trim(SESSION("SECACCESSGRPS"))
    GetCurSecGrp = ""

    strKwGrps = SESSION("KW")
    If left(strKwGrps,1) = "@" Then
       strKwGrps = mid(strKwGrps,2)
    End If

    bolDoPhrases = false
    intSub = 1
    strKwGrp = GetWords(strKwGrps,intSub,1)

    Do While strKwGrp <> ""
       strKwGrp = ReplChars(strKwGrp,CHR(34),"")
       If left(strKwGrp,1) = "-" Then
       ElseIf ucase(strKwGrp) = "OR" or ucase(strKwGrp) = "AND" Then
          strHld = strHld & ","
       Else
          strHld = strHld & " " & strKwGrp
       End If
       intSub = intSub + 1
       strKwGrp = GetWords(strKwGrps,intSub,1)
    Loop
    strKwGrps = trim(strHld)

    bolDoPhrases = true
    intSub = 1
    strSecGrp = GetWords(strSecGrps,intSub,1)

    Do While strSecGrp <> "" and GetCurSecGrp = ""
       intSub = intSub + intRetWords
       intSub2 = 1
       strKwGrp = GetWords(strKwGrps,intSub2,1)
       Do While strKwGrp <> ""
          If ucase(strKwGrp) = ucase(strSecGrp) Then
             GetCurSecGrp = strSecGrp
             EXIT DO
          End If
          intSub2 = intSub2 + intRetWords
          strKwGrp = GetWords(strKwGrps,intSub2,1)
       Loop
       strSecGrp = GetWords(strSecGrps,intSub,1)
    Loop

    If GetCurSecGrp = "" Then
       If ucase(SESSION("SECGRP")) <> "ALL" Then
          GetCurSecGrp = SESSION("SECGRP")
       End If
       If GetCurSecGrp = "" and strKwGrps <> "" Then
          GetCurSecGrp = GetWords(strKwGrps,1,1)
       End If
       If left(GetCurSecGrp,1) = "@" Then
          GetCurSecGrp = mid(GetCurSecGrp,2)
       End If
    End If

END FUNCTION

FUNCTION GetCurPath(GoToLoc)

    If strCurDir = "" Then
       CALL System_Setup("")
    End If
    GetCurPath = "http://" & GetPgmPath(strCurDir)

    If GoToLoc = "Y" Then
      Response.Redirect(GetCurPath)
      Response.END
    End If

END FUNCTION

SUB Process_External_Keywords(SG)

    DIM intSub, strKW, strHld, strWord, strSG

    Session.Contents.Remove("KW")
    Session.Contents.Remove("CAT")
    Session.Contents.Remove("RQL")
    Session.Contents.Remove("MENUFN")

    SESSION("KWO") = "PCN"
    SESSION("RQL") = 1 'Let Logon Group determine access level

    If left(SG,1) = "@" Then
       strSG = mid(SG,2)
       call Read_Text_File(strSG&".DIR",strFileLoc)
       If aryTxtRecs(0) = 0 Then
          If SESSION("$SG") <> "" Then
             strHld = SESSION("$SG")
          ElseIf SESSION("GROUP") <> "" Then
             strHld = SESSION("GROUP")
          ElseIf ucase(SESSION("SECACCESSGRPS")) <> "ALL" Then
             strHld = SESSION("SECACCESSGRPS")
          Else
             strHld = ""
          End If
          aryTxtRecs(1) = "SG:" & strHld
          aryTxtRecs(0) = 1
       End if
    Else
       strSG = SG
       aryTxtRecs(1) = "SG:" & SG
       aryTxtRecs(0) = 1
    End If

    For intSub = 1 to aryTxtRecs(0)
       If trim(aryTxtRecs(intSub)) <> "" Then
          If ucase(left(aryTxtRecs(intSub),3)) = "SG:" Then
             strKW = mid(aryTxtRecs(intSub),4)
          ElseIf ucase(left(aryTxtRecs(intSub),2)) = "G:" Then
             SESSION("GROUP") = trim(mid(aryTxtRecs(intSub),3))
          ElseIf ucase(aryTxtRecs(intSub)) = "N-" Then
             SESSION("KWO") = ReplChars(SESSION("KWO"),"N","") 'Remove Notes Display Option From Menu
          ElseIf ucase(aryTxtRecs(intSub)) = "C-" Then
             SESSION("KWO") = ReplChars(SESSION("KWO"),"C","") 'Remove Company Display Option From Menu
          ElseIf ucase(aryTxtRecs(intSub)) = "N+" Then
             SESSION("KWO") = SESSION("KWO") & "n" 'Include Notes on Default Display
          ElseIf ucase(left(aryTxtRecs(intSub),2)) = "M:" Then
             SESSION("MENUFN") = mid(aryTxtRecs(intSub),3)
          ElseIf left(aryTxtRecs(intSub),1) = "+" Then
             strKW = strKW & " " & mid(aryTxtRecs(intSub),2)
          End If
       End If
    Next
    If instr(ucase(strKW),ucase(strSG)) = 0 Then
       If strKW = "" Then
          strKW = strSG
       Else
          strKW = strSG & "," & strKW
       End If
    End If

    SESSION("KW") = BuildSrchString(strKW)
    If SESSION("KW") <> "" Then
       SESSION("ST") = "SG"
    End If

END SUB

FUNCTION BuildSrchString(KW)

    DIM strKW, intLoc, strHld

    If KW = "" Then

       EXIT FUNCTION
    End If

    strKW = KW & ","

    do while strKW <> ""
       intLoc = instr(strKW,",")
       strHld = trim(left(StrKW,intLoc-1))
       If ucase(strHld) <> "ALL" Then
          If instr(strHld," ") > 0 and left(strHld,1) <> CHR(34) Then
             strHld = CHR(34) & strHld & CHR(34)
          End If
          If BuildSrchString = "" Then
             BuildSrchString = strHld
          Else
             BuildSrchString = BuildSrchString & " OR " & strHld
          End If
       End If
       strKW = trim(mid(strKW,intLoc+1))
    loop

END FUNCTION

%>

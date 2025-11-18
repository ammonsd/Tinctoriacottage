
<!--#INCLUDE FILE="../dir/system_setup.inc" -->
<!--#INCLUDE FILE="../dir/box_it.inc" -->
<!--#INCLUDE FILE="../dir/function_access_groups.inc" -->

<%
    DIM  strSpecSecGrp, bolDisplayMsg, bolSecError

SUB Logon_Check(PgmName,ReqSecLevel,LogonSecGrp)

    DIM strPSWD, strLogonSecGrp, X, strHld

    If strCurDirName = "" Then
       Call System_Setup("NONE")
    End If

    strLogonSecGrp = BuildLogonSecGrp(LogonSecGrp)

    If isnumeric(ReqSecLevel) = false or ReqSecLevel = "" Then
       ReqSecLevel = 9
    Else
       ReqSecLevel = int(ReqSecLevel)
    End If

    If strLogonSecGrp <> "" and strLogonSecGrp <> SESSION("LOGONSECGRP") Then
       SESSION("LOGONSECGRP") = strLogonSecGrp
       Session.Contents.Remove("PSWD")
    End IF

    Session.Contents.Remove("NEWLOGON")

    For Each X In SESSION.CONTENTS
       If trim(SESSION(X)) = "" Then
          Session.Contents.Remove(X)
       End If
    Next

    If SESSION("PSWD") <> "" and SESSION("SECACCESSGRPS") <> "" Then
       If SESSION("SECLEVL") <  ReqSecLevel Then
          If strSpecSecGrp <> "" Then
             bolDisplayMsg = false
             Call Verify_Group_Access("N/R","","")
             If not bolSecError Then
                EXIT SUB
             End If
          End If
          If SESSION("DEBUG") = "Y" Then
             strHld = " / Auth Level: " & ReqSecLevel & " User Level: " & SESSION("SECLEVL")
          Else
             strHld = ""
          End If
          Call Access_Error(PgmName,"User not Authorized for Requested Function" & strHld )
       Else
          If SESSION("KW") = "" Then
            SESSION("KW") = BuildKW("")
          End If
          bolDisplayMsg = true
          CALL Verify_Group_Access(SESSION("KW"),PgmName,ReqSecLevel)
          EXIT SUB
       End If
    End If

    SESSION("P") = BldPgmParms(PgmName)

    SESSION("NEWLOGON") = "Y"
    Session.Contents.Remove("EXPPSWD")

    Response.Redirect(strLogonLoc & "logon.asp")

END SUB

SUB Verify_Group_Access(KW,PGM,SL)

    DIM strWord, intSub, strSecGrp, intSub2, strKwGrps, strAccessGrps, strHld, bolKwFound, intReqSL

    If strSpecSecGrp <> "" Then
       If APPLICATION(strSpecSecGrp) = "N" Then
          bolSecError = true
          EXIT SUB
       End If
    End If

    If SL = "" or not isnumeric(SL) Then
       intReqSL = 9
    Else
       intReqSL = int(SL)
       If intReqSL > 9 Then
          intReqSL = 9
       End If
    End If

    If strSpecSecGrp <> "" and SESSION("SECLEVL") >= intReqSL Then
       strSpecSecGrp = ""
    End If

    If (ucase(SESSION("SECACCESSGRPS")) = "ALL" and strSpecSecGrp = "") OR SESSION("SECLEVL") = 9 Then
       bolSecError = false
       EXIT SUB
    End If

    strKwGrps = ucase(KW)
    If left(strKwGrps,1) = "@" Then
       strKwGrps = mid(strKwGrps,2)
    End If

    bolSecError = true
    bolKwFound = false

    strAccessGrps = trim(ucase(SESSION("SECACCESSGRPS")))

    bolDoPhrases = false
    intSub = 1
    strWord = GetWords(strKwGrps,intSub,1)

    Do While strWord <> ""
       If strWord = "OR" or strWord = "AND" Then
          strHld = strHld & ","
       Else
          strHld = strHld & " " & strWord
       End If
       intSub = intSub + 1
       strWord = GetWords(strKwGrps,intSub,1)
    Loop
    strKwGrps = trim(strHld)
    If strKwGrps = "" Then
       strKwGrps = CHR(254)
    End If

    bolDoPhrases = true
    intSub = 1
    strSecGrp = GetWords(strAccessGrps,intSub,1)

    Do While strSecGrp <> "" and bolSecError
       intSub = intSub + intRetWords
       intSub2 = 1
       strWord = GetWords(strKwGrps,intSub2,1)
       Do While strWord <> ""
          If strSecGrp =  ucase(strSpecSecGrp) Then
             strSpecSecGrp = ""
          End If
          If strWord = strSecGrp or strWord = "N/R" or strSecGrp = "ALL" Then
             bolKwFound = true
          End If
          If bolKwFound and strSpecSecGrp = "" Then
             bolSecError = false
             EXIT DO
          End If
          intSub2 = intSub2 + intRetWords
          strWord = GetWords(strKwGrps,intSub2,1)
       Loop
       strSecGrp = GetWords(strAccessGrps,intSub,1)
    Loop

    If bolSecError and bolDisplayMsg Then
       strHld = ""
       If SESSION("DEBUG") = "Y" and strSpecSecGrp <> "" Then
          strHld = " (" & strSpecSecGrp & ")"
       End If
       CALL Access_Error(PGM,"User not Authorized for Requested Area"  & strHld)
    End If

END SUB

SUB Access_Error(PGM,ERRMSG)

   SESSION("P") = BldPgmParms(PGM)
   CALL Box_Top_Section
   Response.Write "<div align='CENTER'><font color='#BD0000'><b>" & ERRMSG & "</b></font></div>"
   Response.Write "<div align='center'><br><br>"
   Response.Write "<input type='button' value='Back' onClick='history.go(-1)'>&nbsp;&nbsp;&nbsp;"
   Response.Write "<input type='button' value='Logon' onClick=window.location='" & strLogonLoc & "Logon.asp'>"
   Response.Write "</div>"
   CALL Box_Bottom_Section
   Response.END

END SUB

FUNCTION BuildKW(SRCH)

    DIM strWord, intSub, intSub2, strSrch

    strSrch = SRCH
    If strSrch = "" Then
       If SESSION("GROUP") <> "" Then
          strSrch = SESSION("GROUP")
       ElseIf ucase(SESSION("SECACCESSGRPS")) <> "ALL" Then
          strSrch = SESSION("SECACCESSGRPS")
       Else
          strSrch = ""
       End If
    End If

    CALL Get_Function_Security_Groups

    bolDoPhrases = true
    intSub = 1
    strWord = GetWords(strSrch,intSub,1)

    Do While strWord <> ""
       For intSub2 = 1 to aryFunctionGrps(0)
          If ucase(aryFunctionGrps(intSub2)) = ucase(strWord) Then
             strWord = ""
             EXIT FOR
          End If
       Next
       If ucase(strWord) <> "ALL" and strWord <> "" Then
          If BuildKW = "" Then
             BuildKW = strWord
          Else
             BuildKW = BuildKW & " OR " & strWord
          End If
       End If
       intSub = intSub + intRetWords
       strWord = GetWords(strSrch,intSub,1)
    Loop

    SESSION("RQL") = SESSION("SECLEVL")

END FUNCTION

FUNCTION BldPgmParms(PN)

    DIM I, X, strHld, strChkPgm

    If Request.Form.Count >0 THen
       For Each X in Request.Form
          If strHld = "" Then
             strHld = "?" & X & "=" & Request.Form(X)
          Else
             strHld = strHld & "&" & X & "=" & Request.Form(X)
          End If
       Next
    ElseIf Request.QueryString.Count > 0 Then
       For Each X in Request.QueryString
          If X <> "BPG" Then
             If strHld = "" Then
                strHld = "?" & X & "=" & Request.QueryString(X)
             Else
                strHld = strHld & "&" & X & "=" & Request.QueryString(X)
             End If
          End If
       Next
    End If

    BldPgmParms= PN & ReplParmChars(strHld)

    strChkPgm = PN
    For I = len(strChkPgm) to 1 step -1
       if mid(strChkPgm,I,1) = "/" Then
          strChkPgm = mid(strChkPgm,I+1)
          EXIT FOR
       elseif mid(strChkPgm,I,1) = "!" Then
          strChkPgm = mid(strChkPgm,I+1)
          BldPgmParms= left(PN,I-1) & "_" & strChkPgm & ReplParmChars(strHld)
          EXIT FOR
       end if
    Next

    If left(lcase(BldPgmParms),5) <> "http:" Then
       BldPgmParms = GetCurPath("")&BldPgmParms
    End If

END FUNCTION

FUNCTION BuildLogonSecGrp(LogonSecGrp)

    DIM strHld

    If SESSION("LOGONSECGRP") <> "" and lcase(LogonSecGrp) = lcase(SESSION("LOGONSECGRP")) Then
       BuildLogonSecGrp = LogonSecGrp
    Else
       BuildLogonSecGrp = SESSION("HOMEDIR")
       If BuildLogonSecGrp = "" Then
          BuildLogonSecGrp = strCurDirName
       End If
       If LogonSecGrp <> "" Then
          If SESSION("GROUP") <> "" Then
             strHld = "_" & SESSION("GROUP")
          Else
             strHld = ""
          End If
          BuildLogonSecGrp = BuildLogonSecGrp & "_" & LogonSecGrp & strHld
       End If
    End If
    BuildLogonSecGrp = lcase(BuildLogonSecGrp)

END FUNCTION
%>


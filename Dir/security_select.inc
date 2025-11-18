
<!--#INCLUDE FILE="../dir/process_setup.inc" -->
<!--#INCLUDE FILE="../dir/encrypt_decrypt.inc" -->
<!--#INCLUDE FILE="../dir/format_security_groups.inc" -->

<%

    DIM bolBldSecurAryOnly, arySecurity(500,11), bolSecAutoSel, strUserID, strSecGoToPgm
    DIM intSecurityID, strSecFName, strSecLName, strUserName, strPswd, bolExpPswd, strOrigPswd, strSecSecGrps
    DIM intSecLevel, strSecEMail, arySecLevels(10,1), intPswdExpCycle, intUserExpDate, bolLogAccess, bolPwUpdAllowed

FUNCTION SupportDistribution
   call Database_Setup
   SupportDistribution = strSupportMail
END FUNCTION

SUB Security_Selection_List(DefID)

    DIM intSub, intDefID, strDefName, strHld

    Call Database_Setup
    call GoUrl_Code

    If DefID <> "" Then
       DefID = CInt(DefID)
    End If

    Set objRecordSet = Server.CreateObject("ADODB.Recordset")
    objRecordSet.ActiveConnection = strDbConn

    objRecordSet.Source = "SELECT * FROM " & APPLICATION("SECURTBL") & strWhrSQL & " ORDER BY [Last Name]"

    objRecordSet.LockType = adLockReadOnly
    objRecordSet.Open

    intSub = 0
    intDefID = ""

    If not bolBldSecurAryOnly Then
       If bolSecAutoSel then
          Response.Write "<select name='Url' onChange='goUrl(this)'>" & vbCrLf
       Else
          Response.Write "<select name='LID'>" & vbCrLf
       End If
       If DefID = "0" Then
          DefID = ""
          Response.Write "<option value='0'>" & vbCrLf
       End If
    End If

    If Not objRecordSet.EOF Then
       objRecordSet.MoveFirst
       Do While Not objRecordSet.EOF
          intSub = intSub + 1
          arySecurity(intSub,1) = objRecordSet.Fields("ID")
          arySecurity(intSub,2) = trim(objRecordSet.Fields("First Name") & " " & objRecordSet.Fields("Last Name"))
          arySecurity(intSub,3) = objRecordSet.Fields("User Name")
          arySecurity(intSub,4) = objRecordSet.Fields("Email")
          arySecurity(intSub,5) = Decrypt(objRecordSet.Fields("Password"),lcase(objRecordSet.Fields("User Name")))
          arySecurity(intSub,6) = objRecordSet.Fields("Security Level")
          arySecurity(intSub,7) = UnFormatSecurityGroups(objRecordSet.Fields("Security Groups"))
          arySecurity(intSub,8) = objRecordSet.Fields("Last Access")
          arySecurity(intSub,9) = objRecordSet.Fields("User Exp Date")
          arySecurity(intSub,10) = objRecordSet.Fields("First Name")
          arySecurity(intSub,11) = objRecordSet.Fields("Last Name")
          If arySecurity(intSub,1) = DefID Then
             intDefID = arySecurity(intSub,1)
             strDefName = arySecurity(intSub,2)
          End If
          objRecordSet.MoveNext
       Loop
    End If

    arySecurity(0,0) = intSub

    objRecordset.Close
    Set objRecordSet = Nothing

    If bolBldSecurAryOnly Then
       EXIT SUB
    End If

    If intDefID <> "" Then
       Response.Write "<option selected value='" & intDefID & "'>" & strDefName & vbCrLf
    End IF
    For intSub = 1 to arySecurity(0,0)
       If bolSecAutoSel then
          Response.Write "<option value='" & strSecGoToPgm & arySecurity(intSub,1) & "'>" & arySecurity(intSub,2) & " (" & arySecurity(intSub,3) & ")" & "</option>" & vbCrLf
       Else
          Response.Write "<option value='" & arySecurity(intSub,1) & "'>" & arySecurity(intSub,2) & " (" & arySecurity(intSub,3) & ")" & "</option>" & vbCrLf
       End If
    Next
    Response.Write "</select>" & vbCrLf

END SUB

FUNCTION SecurityName(ID)

   DIM intSub

   If isnumeric(ID) = false Then
      EXIT FUNCTION
   END IF

   ID = int(ID)

   Call Build_Security_Array

   For intSub = 1 to arySecLevels(0,0)
      If ID = arySecLevels(intSub,0) Then
        SecurityName = arySecLevels(intSub,1)
        EXIT FUNCTION
      End IF
   Next

END FUNCTION

SUB Security_Level_Selection(CurLevel)

    DIM intSub, strChk, intCntr

    Call Build_Security_Array

    intCntr = 0
    If not isnumeric(CurLevel) Then
       CurLevel = 0
    Else
       CurLevel = int(CurLevel)
    End If

    For intSub = 1 to arySecLevels(0,0)
       If CurLevel = arySecLevels(intSub,0) Then
          strChk = " checked"
       Else
          strChk = ""
       End If
       Response.Write "<input type='radio' name='Security Level' value='" & arySecLevels(intSub,0) & "'" & strChk & ">" & arySecLevels(intSub,1) & vbCrLf
       intCntr = intCntr + 1
       If intCntr = 5 Then
          Response.Write "<br>"  & vbCrLf
          intCntr = 0
       End If
    Next

END SUB

SUB Build_Security_Array

    DIM intSub

    intSub = 1
    arySecLevels(intSub,0) = 9
    arySecLevels(intSub,1) = "Owner"
    intSub = intSub +1
    arySecLevels(intSub,0) = 5
    arySecLevels(intSub,1) = "Administrator"
    intSub = intSub +1
    arySecLevels(intSub,0) = 3
    arySecLevels(intSub,1) = "User"
    intSub = intSub +1
    arySecLevels(intSub,0) = 1
    arySecLevels(intSub,1) = "Access"
    intSub = intSub +1
    arySecLevels(intSub,0) = 0
    arySecLevels(intSub,1) = "None"

    arySecLevels(0,0) = intSub

END SUB

SUB Get_Security(ID)

    Set objRecordSet = Server.CreateObject("ADODB.Recordset")
    objRecordSet.ActiveConnection = strDbConn

    objRecordSet.Source = "SELECT * FROM " & APPLICATION("SECURTBL") & " WHERE [ID] = " & ID

    objRecordSet.LockType = adLockReadOnly
    objRecordSet.Open

    If Not objRecordSet.EOF Then
       objRecordSet.MoveFirst
       intSecurityID = objRecordSet.Fields("ID")
       strSecFName = objRecordSet.Fields("First Name")
       strSecLName = objRecordSet.Fields("Last Name")
       strUserName = lcase(objRecordSet.Fields("User Name"))
       strSecEMail = objRecordSet.Fields("Email")
       strPswd = objRecordSet.Fields("Password")
       If left(strPswd,1) = CHR(254) Then
          bolExpPswd = true
          strPswd = mid(strPswd,2)
       End If
       strPswd = Decrypt(strPswd,strUserName)
       strOrigPswd = strPswd
       intSecLevel = objRecordSet.Fields("Security Level")
       strSecSecGrps = UnFormatSecurityGroups(objRecordSet.Fields("Security Groups"))
       bolLogAccess = objRecordSet.Fields("LogAccess")
       bolPwUpdAllowed = objRecordSet.Fields("Allow Pswd Update")
       intPswdExpCycle = objRecordSet.Fields("Password Exp Cycle")
       intUserExpDate = objRecordSet.Fields("User Exp Date")
    End If

    objRecordset.Close
    Set objRecordSet = Nothing

End SUB

FUNCTION UserNameDuplicate(USERNAME,ID)

    UserNameDuplicate = false

    Call Database_Setup

    Set objRecordSet = Server.CreateObject("ADODB.Recordset")
    objRecordSet.ActiveConnection = strDbConn

    objRecordSet.Source = "SELECT * FROM " & APPLICATION("SECURTBL") & " WHERE [User Name] = '" & lcase(USERNAME) & "'"

    objRecordSet.LockType = adLockReadOnly
    objRecordSet.Open

    If Not objRecordSet.EOF Then
       If objRecordSet.Fields("ID") <> ID Then
          UserNameDuplicate = true
       End If
    End If

    objRecordset.Close
    Set objRecordSet = Nothing

END FUNCTION

SUB Add_Security_Record

   DIM intSub, strWord, strFldName

   If strUserID = "" Then
      strUserID = SESSION("USERID")
   End If

   Set objRecordSet = Server.CreateObject("ADODB.Recordset")
   objRecordSet.Open APPLICATION("SECURTBL"), strDbConn, adOpenStatic, adLockPessimistic, adCmdTable
   on error resume next
   objRecordSet.AddNew

   strFldName = "First Name"
   objRecordSet.Fields(strFldName) = strSecFName

   strFldName = "Last Name"
   objRecordSet.Fields(strFldName) = strSecLName

   strFldName = "User Name"
   objRecordSet.Fields(strFldName) = strUserName

   strFldName = "Email"
   objRecordSet.Fields(strFldName) = strSecEMail

   strFldName = "Password"
   If strPswd = "" Then
      strPswd = "Pass+word"
      bolExpPswd = true
   Else
      bolExpPswd = false
   End If
   objRecordSet.Fields(strFldName) = Encrypt(strPswd,strUserName)

   strFldName = "Security Level"
   Call AddCookie("SECMAINT","SECLEVEL",intSecLevel,365)
   objRecordSet.Fields(strFldName) = intSecLevel

   strFldName = "Security Groups"
   Call AddCookie("SECMAINT","SECGROUPS",strSecSecGrps,365)
   objRecordSet.Fields(strFldName) = FormatSecurityGroups(strSecSecGrps)

   strFldName = "LogAccess"
   If Request.Form(strFldName) = "Y" Then
      bolLogAccess = true
   Else
      bolLogAccess = false
   End If
   Call AddCookie("SECMAINT","LOGACCESS",Request.Form(strFldName),365)
   objRecordSet.Fields(strFldName) = bolLogAccess

   strFldName = "Allow Pswd Update"
   If Request.Form(strFldName) = "Y" Then
      bolPwUpdAllowed = true
   Else
      bolPwUpdAllowed = false
   End If
   Call AddCookie("SECMAINT","PSWDUPDOK",Request.Form(strFldName),365)
   objRecordSet.Fields(strFldName) = bolPwUpdAllowed

   If bolExpPswd Then
      objRecordSet.Fields("Password Exp Date") = dateadd("d", -1, Date)
   Else
      objRecordSet.Fields("Password Exp Date") = dateadd("d", int(intPswdExpCycle), Date)
   End If

   strFldName = "Password Exp Cycle"
   Call AddCookie("SECMAINT","EXPPSWD",intPswdExpCycle,365)
   objRecordSet.Fields(strFldName) = intPswdExpCycle

   strFldName = "User Exp Date"
   If intUserExpDate = "" Then
      intUserExpDate = "12/31/9999"
   End If
   intUserExpDate = CDate(intUserExpDate)
   Call AddCookie("SECMAINT","EXPDATE",intUserExpDate,365)
   objRecordSet.Fields(strFldName) = intUserExpDate

   objRecordSet.Fields("Added") = NOW()
   objRecordSet.Fields("Added User ID") = strUserID
   objRecordSet.Fields("Last Access") = NOW()
   objRecordSet.Fields("Logon Attempts") = 0

   objRecordSet.Update

   objRecordSet.MoveLast
   intSecurityID = objRecordSet.Fields("ID")

   objRecordSet.Close

   If instr(strSecSecGrps,",") > 0 Then
      strSecSecGrps = left(strSecSecGrps,instr(strSecSecGrps,",")-1)
   End If

   Call Security_Notification("A",strSecSecGrps)

END SUB

SUB Validate_Security_Entry

  '// Validate form entries.

   DIM intSub

   intErrNbr = 0

   If trim(FormData("First Name")) = "" and trim(FormData("Last Name")) = "" Then
      intErrNbr = intErrNbr + 1
      errMsgs.Add intErrNbr, "Either a First or Last Name Must Be Entered"
   End If
   If trim(strUserName) = "" Then
      intErrNbr = intErrNbr + 1
      errMsgs.Add intErrNbr, "User Name Required"
   ElseIf len(trim(strUserName)) < 4 Then
      intErrNbr = intErrNbr + 1
      errMsgs.Add intErrNbr, "User Name Must be At Least four Characters"
   ElseIf strUserName <> Request.Form("OrigUName") Then
      If UserNameDuplicate(strUserName,intSecurityID) then
         intErrNbr = intErrNbr + 1
         errMsgs.Add intErrNbr, "User Name Already Assigned"
      End If
   End If
   If trim(intSecLevel) = "" Then
      intErrNbr = intErrNbr + 1
      errMsgs.Add intErrNbr, "Security Level Required"
   End If
   If strSecEMail = "" then
      intErrNbr = intErrNbr + 1
      errMsgs.Add intErrNbr, "Email Address Required"
   ElseIf not ValidateEMailFormat(strSecEMail) then
      intErrNbr = intErrNbr + 1
      errMsgs.Add intErrNbr, "Invalid Email Address Format"
   End If

   If Request.Form("Allow Pswd Update") = "Y" Then
      If not isnumeric(Request.Form("Password Exp Cycle")) or Request.Form("Password Exp Cycle") < "1" Then
         intErrNbr = intErrNbr + 1
         errMsgs.Add intErrNbr, "Password Expiration Cycle must be greater than zero days"
      Else
         intPswdExpCycle = Request.Form("Password Exp Cycle")
      End If
   Else
      intPswdExpCycle = 9999
   End If
   If intUserExpDate <> "" Then
      If instr(intUserExpDate,"/") > 0 Then
         If not isdate(intUserExpDate) Then
            intErrNbr = intErrNbr + 1
            errMsgs.Add intErrNbr, "Format of User Expiration Date is Invalid (MM/DD/YYYY)"
         End If
      ElseIf not isnumeric(intUserExpDate) Then
         intErrNbr = intErrNbr + 1
         errMsgs.Add intErrNbr, "Number of Days Until User Expiration Date is Not Numeric"
      End If
   End If

   If strSecSecGrps = "" Then
      strSecSecGrps = "All"
   Else
      For intSub = 1 to aryFunctionGrps(0)
         If ucase(aryFunctionGrps(intSub)) = ucase(strSecSecGrps) Then
            strSecSecGrps = "All, " & strSecSecGrps
            EXIT FOR
         End If
      Next
   End If

End SUB
%>


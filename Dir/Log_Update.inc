
<%

SUB Update_Access_Log(PGM)

    DIM strFn, strPgm

    strPgm = PGM
    strFn = SESSION("HOMEDIR")
    If strFn = "" Then
       strFn = strCurDir
    End If
    If strFn <> "" Then
       strFn = strFn & "_"
    End If

    strFn = strFn & strLogFn

    If strPgm <> "" Then
       strPgm = " - " & strPgm
    End If

    If SESSION("SECLOG") = true Then
       bolAppendFile = true
       aryTxtRecs(1) = NOW & ": " & SESSION("USERNAME") & " (" & SESSION("USEREMAIL") & ")" & strPgm
       aryTxtRecs(0) = 1
       call Write_Text_File(strFn,strFileLoc)
    End If

END SUB
%>

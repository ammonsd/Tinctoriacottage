<%FUNCTION GetPgmPath(strDir)

   DIM  intLoc

   GetPgmPath = strDir

   If GetPgmPath = "" Then
      GetPgmPath = Request.ServerVariables("SCRIPT_NAME")
      For intLoc = len(GetPgmPath) to 1 step -1
         If mid(GetPgmPath,intLoc,1) = "/" or mid(GetPgmPath,intLoc,1) = "\" Then
            GetPgmPath = left(GetPgmPath,intLoc)
            Exit For
         End If
      Next
   ElseIf GetPgmPath <> "/" Then
      GetPgmPath = "/" & GetPgmPath & "/"
   End If
   GetPgmPath = Request.ServerVariables("SERVER_NAME") & GetPgmPath
END FUNCTION %>

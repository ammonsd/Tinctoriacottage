
<%

   DIM aryText(15,1)

FUNCTION CheckTextVariables(strMsg)

   DIM intLoc, intSub

   aryText(0,0) = 15

   For intSub = 1 to aryText(0,0)
      If aryText(intSub,0) = "" Then
         aryText(intSub,0) = "[var" & intSub & "]"
      End If
   Next

   CheckTextVariables = strMsg

   For intSub = 1 to aryText(0,0)
      Do
         If instr(CheckTextVariables,"[") = 0 and instr(CheckTextVariables,"]") = 0 Then
            Exit FOR
         End If
         intLoc = instr(CheckTextVariables,aryText(intSub,0))
         If intLoc > 0 Then
            CheckTextVariables = left(CheckTextVariables,intLoc-1) & aryText(intSub,1) & mid(CheckTextVariables,intLoc + len(aryText(intSub,0)))
         End If
      Loop Until intLoc = 0
   Next

END FUNCTION%>

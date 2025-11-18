<%
FUNCTION CheckNumeric(NumString,AddlChkString)

 'Check for only numerics and any characters in AddlChkString (i.e. a - or ."

   DIM strChk, intCntr, strNumString

   strNumString = trim(NumString)
   strChk = "0123456789" & AddlChkString
   CheckNumeric = True

   For intCntr = 1 to len(strNumString)
      If instr(strChk,mid(strNumString,intCntr,1)) = 0 Then
         CheckNumeric = False
         Exit For
      End If
   Next

End FUNCTION
%>

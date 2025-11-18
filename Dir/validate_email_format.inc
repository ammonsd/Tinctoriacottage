
<% FUNCTION ValidateEMailFormat(EMAIL)

   DIM strValidChars, intCntr

   ValidateEMailFormat = true

   intCntr = instr(EMAIL,"@")

   If intCntr = 0 Then
      ValidateEMailFormat = false
   ElseIf instr(intCntr,EMAIL,".") = 0 Then
      ValidateEMailFormat = false
   Else
      strValidChars  = "0123456789._-@"
      For intCntr = 65 to 90
         strValidChars  = strValidChars  & CHR(intCntr)
      Next
      For intCntr = 1 To len(EMAIL)
         If instr(strValidChars ,mid(UCase(EMAIL),intCntr,1)) = 0 Then
            ValidateEMailFormat = false
            EXIT FUNCTION
         End If
      Next
   End If

END FUNCTION %>

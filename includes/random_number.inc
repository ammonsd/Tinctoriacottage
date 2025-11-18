
<%  FUNCTION RandomNbr(intSize)
  
   DIM I
   
   Randomize()
   RandomNbr = Rnd()  
   
   If isNumeric(intSize) = true Then
      RandomNbr = mid(RandomNbr,instr(RandomNbr,".")+1,intSize)
   Else
      RandomNbr = mid(RandomNbr,instr(RandomNbr,".")+1)
   End If
   
End FUNCTION %>

<%
   DIM intNbrEntries, arySorted

FUNCTION SortArray(arySort, intAsc)

   DIM strHld, intSub, intSub2, bolSwap

   For intSub = 0 To UBound(arySort) - 1
      If trim(arySort(intSub)) <> "" and not isnull(arySort(intSub)) Then
         For intSub2 = 0 To UBound(arySort)
            bolSwap = false
            'Sort Ascending
            if intAsc = 1 Then
               If isNumeric(arySort(intSub)) and isNumeric(arySort(intSub2)) Then
                  if int(arySort(intSub)) > int(arySort(intSub2)) Then
                     bolSwap = true
                  End if
               Else
                  if lcase(arySort(intSub)) > lcase(arySort(intSub2)) Then
                     bolSwap = true
                  End if
               End if
            'Sort Descending
            Else
               If isNumeric(arySort(intSub)) and isNumeric(arySort(intSub2)) Then
                  if int(arySort(intSub)) < int(arySort(intSub2)) Then
                     bolSwap = true
                  End if
               Else
                  if lcase(arySort(intSub)) < lcase(arySort(intSub2)) Then
                     bolSwap = true
                  End if
               End if
            End if
            If bolSwap Then
               strHld = arySort(intSub)
               arySort(intSub) = arySort(intSub2)
               arySort(intSub2) = strHld
            End If
         Next
      End If
   Next

   intNbrEntries = 0
   For intSub = 0 To UBound(arySort) - 1
      If trim(arySort(intSub)) = "" or isnull(arySort(intSub)) Then
         EXIT FOR
      End If
      intNbrEntries = intNbrEntries + 1
   Next

   SortArray = arySort
End FUNCTION

FUNCTION SortArray2(arySort, intAsc)

   DIM strHld, intSub, intSub2, intSub3, bolSwap

   For intSub = 0 To UBound(arySort,1) - 1
      If trim(arySort(intSub,1)) <> "" and not isnull(arySort(intSub,1)) Then
         For intSub2 = 0 To UBound(arySort,1)
            bolSwap = false
            'Sort Ascending
            if intAsc = 1 Then
               If isNumeric(arySort(intSub,1)) and isNumeric(arySort(intSub2,1)) Then
                  if int(arySort(intSub,1)) > int(arySort(intSub2,1)) Then
                     bolSwap = true
                  End if
               Else
                  if lcase(arySort(intSub,1)) > lcase(arySort(intSub2,1)) Then
                     bolSwap = true
                  End if
               End if
            'Sort Descending
            Else
               If isNumeric(arySort(intSub,1)) and isNumeric(arySort(intSub2,1)) Then
                  if int(arySort(intSub,1)) < int(arySort(intSub2,1)) Then
                     bolSwap = true
                  End if
               Else
                  if lcase(arySort(intSub,1)) < lcase(arySort(intSub2,1)) Then
                     bolSwap = true
                  End if
               End if
            End if
            If bolSwap Then
               For intSub3 = 0 to UBound(arySort,2)
                  strHld = arySort(intSub,intSub3)
                  arySort(intSub,intSub3) = arySort(intSub2,intSub3)
                  arySort(intSub2,intSub3) = strHld
               Next
            End If
         Next
      End If
   Next

   intNbrEntries = 0
   For intSub = 0 To UBound(arySort,1) - 1
      If trim(arySort(intSub,1)) <> "" and not isnull(arySort(intSub,1)) Then
         intNbrEntries = intNbrEntries + 1
      End If
   Next

   SortArray2 = arySort
End FUNCTION
%>


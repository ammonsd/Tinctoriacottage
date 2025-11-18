
<!--#INCLUDE FILE="../dir/member_select.inc" -->

<%

   DIM bolGetMbrDetails

SUB Load_Variables(EM)

   DIM strHld, intID, intSub, intLoc

   bolDoPhrases = false
   strWrdSep = ""

   strHld = EM
   For intSub = len(strHld) to 1 step -1
      If mid(strHld,intSub,1) = " " Then
         strEmail = mid(strHld,intSub+1)
         strEmail = ReplChars(strEmail,"<"," ")
         strEmail = trim(ReplChars(strEmail,">"," "))
         strHld = trim(left(strHld,intSub-1))
         intLoc=TotalWords(strHld)
         If intLoc=1 Then
            strLName = strHld
         Else
            strLName = GetWords(strHld,intLoc,1)
            strFName = GetWords(strHld,1,intLoc-1)
         End If
         EXIT FOR
      End If
   Next

   aryText(1,0) = "[ToFName]"
   aryText(1,1) = strFName
   aryText(2,0) = "[ToLName]"
   aryText(2,1) = strLName
   aryText(3,0) = "[SenderName]"
   aryText(3,1) = strFrName
   aryText(4,0) = "[SenderEmail]"
   aryText(4,1) = strFrEmail
   aryText(5,0) = "[Email]"
   aryText(5,1) = strEmail

   intSub = 5
   If bolGetMbrDetails Then
      intID = GetMbrID(strEmail)
      Call Get_DB_Record(intID)
      intSub = intSub + 1
      aryText(intSub,0) = "[Street]"
      If strStreet = "" and strCity = "" and strState = "" and strZip = "" Then
         aryText(intSub,1) = strCountry
         strCountry = ""
      Else
         aryText(intSub,1) = strStreet
      End If
      intSub = intSub + 1
      aryText(intSub,0) = "[City]"
      aryText(intSub,1) = strCity
      intSub = intSub + 1
      aryText(intSub,0) = "[State]"
      aryText(intSub,1) = strState
      intSub = intSub + 1
      aryText(intSub,0) = "[Zip]"
      aryText(intSub,1) = strZip
      intSub = intSub + 1
      aryText(intSub,0) = "[CityStateZip]"
      aryText(intSub,1) = strCity
      If strState <> "" Then
         If strCity <> "" Then
            aryText(intSub,1) = aryText(intSub,1) & ", " & strState
         Else
            aryText(intSub,1) = strState
         End If
      End If
      aryText(intSub,1) = aryText(intSub,1) & " " & strZip
      intSub = intSub + 1
      aryText(intSub,0) = "[Country]"
      aryText(intSub,1) = strCountry
      intSub = intSub + 1
      aryText(intSub,0) = "[LastUpdate]"
      If strDateUpdated <> "" Then
         aryText(intSub,1) = strDateUpdated
      Else
         aryText(intSub,1) = strDateAdded
      End If
      aryText(intSub,1) = GetWords(aryText(intSub,1),1,1)
    End If

    aryText(0,0) = intSub

END SUB
%>

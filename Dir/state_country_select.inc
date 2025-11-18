
<%

   DIM arySelData(100,2), bolSelAryOnly, bolStateReq, bolCountryReq, intBlnkEntry

SUB Build_States_Select_List(CUR)

   DIM intSub, strCurV, strCurN, intFirstEntry

   intSub = 0
   If not bolStateReq Then
      intSub = intSub + 1
      arySelData(intSub,1) = ""
      arySelData(intSub,2) = ""
   End If
   intBlnkEntry = intSub
   intSub = intSub + 1
   arySelData(intSub,1) = "AL"
   arySelData(intSub,2) = "Alabama"
   intSub = intSub + 1
   arySelData(intSub,1) = "AK"
   arySelData(intSub,2) = "Alaska"
   intSub = intSub + 1
   arySelData(intSub,1) = "AZ"
   arySelData(intSub,2) = "Arizona"
   intSub = intSub + 1
   arySelData(intSub,1) = "AR"
   arySelData(intSub,2) = "Arkansas"
   intSub = intSub + 1
   arySelData(intSub,1) = "CA"
   arySelData(intSub,2) = "California"
   intSub = intSub + 1
   arySelData(intSub,1) = "CO"
   arySelData(intSub,2) = "Colorado"
   intSub = intSub + 1
   arySelData(intSub,1) = "CT"
   arySelData(intSub,2) = "Connecticut"
   intSub = intSub + 1
   arySelData(intSub,1) = "DE"
   arySelData(intSub,2) = "Delaware"
   intSub = intSub + 1
   arySelData(intSub,1) = "DC"
   arySelData(intSub,2) = "District of Columbia"
   intSub = intSub + 1
   arySelData(intSub,1) = "FL"
   arySelData(intSub,2) = "Florida"
   intSub = intSub + 1
   arySelData(intSub,1) = "GA"
   arySelData(intSub,2) = "Georgia"
   intSub = intSub + 1
   arySelData(intSub,1) = "HI"
   arySelData(intSub,2) = "Hawaii"
   intSub = intSub + 1
   arySelData(intSub,1) = "ID"
   arySelData(intSub,2) = "Idaho"
   intSub = intSub + 1
   arySelData(intSub,1) = "IL"
   arySelData(intSub,2) = "Illinois"
   intSub = intSub + 1
   arySelData(intSub,1) = "IN"
   arySelData(intSub,2) = "Indiana"
   intSub = intSub + 1
   arySelData(intSub,1) = "IA"
   arySelData(intSub,2) = "Iowa"
   intSub = intSub + 1
   arySelData(intSub,1) = "KS"
   arySelData(intSub,2) = "Kansas"
   intSub = intSub + 1
   arySelData(intSub,1) = "KY"
   arySelData(intSub,2) = "Kentucky"
   intSub = intSub + 1
   arySelData(intSub,1) = "LA"
   arySelData(intSub,2) = "Louisiana"
   intSub = intSub + 1
   arySelData(intSub,1) = "ME"
   arySelData(intSub,2) = "Maine"
   intSub = intSub + 1
   arySelData(intSub,1) = "MD"
   arySelData(intSub,2) = "Maryland"
   intSub = intSub + 1
   arySelData(intSub,1) = "MA"
   arySelData(intSub,2) = "Massachusetts"
   intSub = intSub + 1
   arySelData(intSub,1) = "MI"
   arySelData(intSub,2) = "Michigan"
   intSub = intSub + 1
   arySelData(intSub,1) = "MN"
   arySelData(intSub,2) = "Minnesota"
   intSub = intSub + 1
   arySelData(intSub,1) = "MS"
   arySelData(intSub,2) = "Mississippi"
   intSub = intSub + 1
   arySelData(intSub,1) = "MO"
   arySelData(intSub,2) = "Missouri"
   intSub = intSub + 1
   arySelData(intSub,1) = "MT"
   arySelData(intSub,2) = "Montana"
   intSub = intSub + 1
   arySelData(intSub,1) = "NE"
   arySelData(intSub,2) = "Nebraska"
   intSub = intSub + 1
   arySelData(intSub,1) = "NV"
   arySelData(intSub,2) = "Nevada"
   intSub = intSub + 1
   arySelData(intSub,1) = "NH"
   arySelData(intSub,2) = "New Hampshire"
   intSub = intSub + 1
   arySelData(intSub,1) = "NJ"
   arySelData(intSub,2) = "New Jersey"
   intSub = intSub + 1
   arySelData(intSub,1) = "NM"
   arySelData(intSub,2) = "New Mexico"
   intSub = intSub + 1
   arySelData(intSub,1) = "NY"
   arySelData(intSub,2) = "New York"
   intSub = intSub + 1
   arySelData(intSub,1) = "NC"
   arySelData(intSub,2) = "North Carolina"
   intSub = intSub + 1
   arySelData(intSub,1) = "ND"
   arySelData(intSub,2) = "North Dakota"
   intSub = intSub + 1
   arySelData(intSub,1) = "OH"
   arySelData(intSub,2) = "Ohio"
   intSub = intSub + 1
   arySelData(intSub,1) = "OK"
   arySelData(intSub,2) = "Oklahoma"
   intSub = intSub + 1
   arySelData(intSub,1) = "OR"
   arySelData(intSub,2) = "Oregon"
   intSub = intSub + 1
   arySelData(intSub,1) = "PA"
   arySelData(intSub,2) = "Pennsylvania"
   intSub = intSub + 1
   arySelData(intSub,1) = "PR"
   arySelData(intSub,2) = "Puerto Rico"
   intSub = intSub + 1
   arySelData(intSub,1) = "RI"
   arySelData(intSub,2) = "Rhode Island"
   intSub = intSub + 1
   arySelData(intSub,1) = "SC"
   arySelData(intSub,2) = "South Carolina"
   intSub = intSub + 1
   arySelData(intSub,1) = "SD"
   arySelData(intSub,2) = "South Dakota"
   intSub = intSub + 1
   arySelData(intSub,1) = "TN"
   arySelData(intSub,2) = "Tennessee"
   intSub = intSub + 1
   arySelData(intSub,1) = "TX"
   arySelData(intSub,2) = "Texas"
   intSub = intSub + 1
   arySelData(intSub,1) = "UT"
   arySelData(intSub,2) = "Utah"
   intSub = intSub + 1
   arySelData(intSub,1) = "VT"
   arySelData(intSub,2) = "Vermont"
   intSub = intSub + 1
   arySelData(intSub,1) = "VA"
   arySelData(intSub,2) = "Virgina"
   intSub = intSub + 1
   arySelData(intSub,1) = "WA"
   arySelData(intSub,2) = "Washington"
   intSub = intSub + 1
   arySelData(intSub,1) = "WV"
   arySelData(intSub,2) = "West Virgina"
   intSub = intSub + 1
   arySelData(intSub,1) = "WI"
   arySelData(intSub,2) = "Wisconson"
   intSub = intSub + 1
   arySelData(intSub,1) = "WY"
   arySelData(intSub,2) = "Wyoming"
   arySelData(0,0) = intSub

   If bolSelAryOnly Then
      EXIT SUB
   End If

   If CUR <> "" Then
      For intSub = intBlnkEntry+1 to arySelData(0,0)
         If arySelData(intSub,1) = CUR Then
            strCurV = arySelData(intSub,1)
            strCurN = arySelData(intSub,2)
            arySelData(intSub,1) = ""
            EXIT FOR
         End If
      Next
   End If
   Response.Write "<select name='State' size='1'>" & vbCrLf
   Response.Write "<option selected value='" & strCurV & "'>" & strCurN & "</option>" & vbCrLf
   For intSub = 1 to arySelData(0,0)
     If arySelData(intSub,1) <> "" or intSub = intBlnkEntry Then
        Response.Write "<option value='" & arySelData(intSub,1) & "'>" & arySelData(intSub,2) & "</option>" & vbCrLf
     End If
   Next
   Response.Write "</select>" & vbCrLf
END SUB

SUB Build_States_Select_List2(CUR)

   DIM intSub, strCurV, strCurN

   bolSelAryOnly = true
   Call Build_States_Select_List("")

   If CUR <> "" Then
      For intSub = intBlnkEntry+1 to arySelData(0,0)
         If arySelData(intSub,1) = CUR Then
            strCurV = arySelData(intSub,1)
            strCurN = arySelData(intSub,1)
            arySelData(intSub,1) = ""
            EXIT FOR
         End If
      Next
   End If
   Response.Write "<select name='State' size='1'>" & vbCrLf
   Response.Write "<option selected value='" & strCurV & "'>" & strCurN & "</option>" & vbCrLf
   For intSub = 1 to arySelData(0,0)
     If arySelData(intSub,1) <> "" or intSub = intBlnkEntry Then
        Response.Write "<option value='" & arySelData(intSub,1) & "'>" & arySelData(intSub,1) & "</option>" & vbCrLf
     End If
   Next
   Response.Write "</select>" & vbCrLf
END SUB

SUB Build_Country_Select_List(CUR)

   DIM intSub, strCurV, strCur

   strCur = CUR

   intSub = 0
   If not bolCountryReq Then
      intSub = intSub + 1
      arySelData(intSub,1) = ""
      arySelData(intSub,2) = ""
   End If
   intBlnkEntry = intSub
   arySelData(intSub,1) = ""
   arySelData(intSub,2) = ""
   intSub = intSub + 1
   arySelData(intSub,1) = "Angola"
   intSub = intSub + 1
   arySelData(intSub,1) = "Argentina"
   intSub = intSub + 1
   arySelData(intSub,1) = "Austria"
   intSub = intSub + 1
   arySelData(intSub,1) = "Australia"
   intSub = intSub + 1
   arySelData(intSub,1) = "Belgium"
   intSub = intSub + 1
   arySelData(intSub,1) = "Bolivia"
   intSub = intSub + 1
   arySelData(intSub,1) = "Brazil"
   intSub = intSub + 1
   arySelData(intSub,1) = "Canada"
   intSub = intSub + 1
   arySelData(intSub,1) = "Chile"
   intSub = intSub + 1
   arySelData(intSub,1) = "China"
   intSub = intSub + 1
   arySelData(intSub,1) = "Columbia"
   intSub = intSub + 1
   arySelData(intSub,1) = "Czech Republic"
   intSub = intSub + 1
   arySelData(intSub,1) = "Denmark"
   intSub = intSub + 1
   arySelData(intSub,1) = "Egypt"
   intSub = intSub + 1
   arySelData(intSub,1) = "England"
   intSub = intSub + 1
   arySelData(intSub,1) = "Estonia"
   intSub = intSub + 1
   arySelData(intSub,1) = "Finland"
   intSub = intSub + 1
   arySelData(intSub,1) = "France"
   intSub = intSub + 1
   arySelData(intSub,1) = "Ghana"
   intSub = intSub + 1
   arySelData(intSub,1) = "Germany"
   intSub = intSub + 1
   arySelData(intSub,1) = "Greece"
   intSub = intSub + 1
   arySelData(intSub,1) = "Hong Kong"
   intSub = intSub + 1
   arySelData(intSub,1) = "Hungary"
   intSub = intSub + 1
   arySelData(intSub,1) = "Iceland"
   intSub = intSub + 1
   arySelData(intSub,1) = "India"
   intSub = intSub + 1
   arySelData(intSub,1) = "Indonesia"
   intSub = intSub + 1
   arySelData(intSub,1) = "Iran"
   intSub = intSub + 1
   arySelData(intSub,1) = "Ireland"
   intSub = intSub + 1
   arySelData(intSub,1) = "Israel"
   intSub = intSub + 1
   arySelData(intSub,1) = "Italy"
   intSub = intSub + 1
   arySelData(intSub,1) = "Jamaica"
   intSub = intSub + 1
   arySelData(intSub,1) = "Japan"
   intSub = intSub + 1
   arySelData(intSub,1) = "Korea"
   intSub = intSub + 1
   arySelData(intSub,1) = "Latvia"
   intSub = intSub + 1
   arySelData(intSub,1) = "Lithuania"
   intSub = intSub + 1
   arySelData(intSub,1) = "Malaysia"
   intSub = intSub + 1
   arySelData(intSub,1) = "Mexico"
   intSub = intSub + 1
   arySelData(intSub,1) = "Micronesia"
   intSub = intSub + 1
   arySelData(intSub,1) = "Nambia"
   intSub = intSub + 1
   arySelData(intSub,1) = "New Zealand"
   intSub = intSub + 1
   arySelData(intSub,1) = "Nigeria"
   intSub = intSub + 1
   arySelData(intSub,1) = "Norway"
   intSub = intSub + 1
   arySelData(intSub,1) = "Paraguay"
   intSub = intSub + 1
   arySelData(intSub,1) = "Philippines"
   intSub = intSub + 1
   arySelData(intSub,1) = "Poland"
   intSub = intSub + 1
   arySelData(intSub,1) = "Portugal"
   intSub = intSub + 1
   arySelData(intSub,1) = "Puerto Rico"
   intSub = intSub + 1
   arySelData(intSub,1) = "Romania"
   intSub = intSub + 1
   arySelData(intSub,1) = "Russia"
   intSub = intSub + 1
   arySelData(intSub,1) = "Saudi Arabia"
   intSub = intSub + 1
   arySelData(intSub,1) = "Scotland"
   intSub = intSub + 1
   arySelData(intSub,1) = "Seychelles"
   intSub = intSub + 1
   arySelData(intSub,1) = "Singapore"
   intSub = intSub + 1
   arySelData(intSub,1) = "South Africa"
   intSub = intSub + 1
   arySelData(intSub,1) = "Spain"
   intSub = intSub + 1
   arySelData(intSub,1) = "Sweden"
   intSub = intSub + 1
   arySelData(intSub,1) = "Switzerland"
   intSub = intSub + 1
   arySelData(intSub,1) = "Tabago"
   intSub = intSub + 1
   arySelData(intSub,1) = "Taiwan"
   intSub = intSub + 1
   arySelData(intSub,1) = "The Netherlands"
   intSub = intSub + 1
   arySelData(intSub,1) = "Thialand"
   intSub = intSub + 1
   arySelData(intSub,1) = "Trinidad"
   intSub = intSub + 1
   arySelData(intSub,1) = "Turkey"
   intSub = intSub + 1
   arySelData(intSub,1) = "UAE"
   intSub = intSub + 1
   arySelData(intSub,1) = "United Kingdom"
   intSub = intSub + 1
   arySelData(intSub,1) = "United States"
   If strCur = "" Then
      strCur = arySelData(intSub,1)
   End If
   intSub = intSub + 1
   arySelData(intSub,1) = "Uruguay"
   intSub = intSub + 1
   arySelData(intSub,1) = "Venezuela"
   intSub = intSub + 1
   arySelData(intSub,1) = "Zaire"
   intSub = intSub + 1
   arySelData(intSub,1) = "Zimbabwe"
   arySelData(0,0) = intSub

   For intSub = intBlnkEntry+1 to arySelData(0,0)
      If arySelData(intSub,1) = strCur Then
         strCurV = arySelData(intSub,1)
         arySelData(intSub,1) = ""
         EXIT FOR
      End If
   Next

   Response.Write "<select name='Country' size='1'>" & vbCrLf
   Response.Write "<option selected value='" & strCurV & "'>" & strCurV & "</option>" & vbCrLf
   For intSub = 1 to arySelData(0,0)
     If arySelData(intSub,1) <> "" or intSub = intBlnkEntry Then
        Response.Write "<option value='" & arySelData(intSub,1) & "'>" & arySelData(intSub,1) & "</option>" & vbCrLf
     End If
   Next
   Response.Write "</select>" & vbCrLf

END SUB
%>

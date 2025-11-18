<%

	
	If Request.QueryString("C") <> "" Then 
	   SESSION("C") = Request.QueryString("C")
	End If
	
FUNCTION SetLogonGrp(PgmName,UsrGrps)
	
	DIM strSel
	
	strSel = ucase(PgmName)
	If instr(strSel,".") > 0 Then
	   strSel = left(strSel,instr(strSel,".") - 1)
	End If
	SESSION("AccessLvl") = "R"
	If SESSION("MA") = "Y" Then
	   SetLogonGrp = "@"
	   EXIT FUNCTION
	End If
	Select CASE SESSION("C")
	  CASE "", "1"
	    If instr(UsrGrps," INVEN ") > 0 Then
		   Select CASE strSel
		      CASE "BOOKPG_MAINT","WOVENPG_MAINT","CREATIONPG_MAINT","LINENPG_MAINT"
		         SetLogonGrp = "INVEN"
		         SESSION("AccessLvl") = "U"
		   End Select
	    End If
	    If instr(UsrGrps," READONLY ") > 0 Then
		   Select CASE strSel
		      CASE "MEMBERS_SEL","EMAIL_LIST","LIST_MEMBERS","MEMBERS_UPDATE"
		         SetLogonGrp = "READONLY"
		         SESSION("AccessLvl") = "R"
		   End Select
	    End If
	END Select
END FUNCTION%>

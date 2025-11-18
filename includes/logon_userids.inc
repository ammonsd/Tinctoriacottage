<%
	DIM aryLogonInfo()
	
	If Request.QueryString("C") <> "" Then 
	   SESSION("C") = Request.QueryString("C")
	End If
	
	Select CASE SESSION("C")
	  CASE "", "1"
		ReDim aryLogonInfo(2,3)
		aryLogonInfo(0,0) = 2
		aryLogonInfo(1,1) = "AMMONSD"
		aryLogonInfo(1,2) = "‹ ™ˆ”[_›Žˆ"
		aryLogonInfo(1,3) = " ADMIN "
		aryLogonInfo(2,1) = "AMMONST"
		aryLogonInfo(2,2) = "‹ ™ˆ”WX›Žˆ"
		aryLogonInfo(2,3) = " INVEN "
	END Select
%>

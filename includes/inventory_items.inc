<%
DIM aryItems()
If Request.QueryString("C") <> "" Then 
   SESSION("C") = Request.QueryString("C")
End If
FUNCTION GetInventory
Select CASE SESSION("C")
CASE "", "1"
ReDim aryItems(6,3)
aryItems(0,0) = 6
aryItems(1,1) = "Handwoven Cotton Towels"
aryItems(1,2) = "HW1001"
aryItems(1,3) = "12.5"
aryItems(2,1) = "Handwoven Blanket, Small"
aryItems(2,2) = "HW1002"
aryItems(2,3) = "55"
aryItems(3,1) = "Handwoven Cotton Towel"
aryItems(3,2) = "HW1003"
aryItems(3,3) = "14"
aryItems(4,1) = "Handmade Decorative Pillow"
aryItems(4,2) = "VC1001"
aryItems(4,3) = "24"
aryItems(5,1) = "Handmade Decorative Pillow"
aryItems(5,2) = "VC1002"
aryItems(5,3) = "28"
aryItems(6,1) = "Vintage Textile Printing Block"
aryItems(6,2) = "VT1001"
aryItems(6,3) = "25"
END Select
END FUNCTION
%>

<!--#INCLUDE FILE="../includes/company_setup_common.inc" -->
		
<%SUB Setup_Company_Info
		
	Select CASE SESSION("C")
	  CASE "", "1"
		
		strCompanyName = "Tinctoria Cottage"
		strCompanyAddr1 = "1 Good Hope Road"
		strCompanyAddr2 = "Elkton, MD 21921"
		strCompanyPhone = "410.398.9326"
		strCompanyFax= "410.398.9643"
		strCompanyHomePg = "www.tinctoriacottage.com"
		strSalesTaxStates = "MD"
		strSalesTaxMsg = "MD Residents add 5% Sales Tax"
		strShippingDesc = "$3.50 for the first item, $0.50 for each additional item."
		strCreditCardMsg = "We do not currently accept credit cards."
		strReturnChkMsg = "* Checks returned for non-sufficient funds (NSF), will be assessed a $20.00."
		strCustSrveMail = "CustomerService@tinctoriacottage.com"
		strOrderDepteMail = "OrderDept@tinctoriacottage.com"
		strSMTPServer = "mail.tinctoriacottage.com"
		strFormClr = "#FFFFCC"
		strCompanyWinClr = " bgcolor = '#FFFFCC'"
		strUpdateDB = "true"
		strCompanyNotify = "true"
		strCommentsDisplay = "true"
		strEmailList = "true"
		strBCastEMail = "true"
		strCompanyDisplay = "false"
		strStreetDisplay = "true"
		strAddress2Display = "false"
		strCityDisplay = "true"
		strStateDisplay = "true"
		strZipDisplay = "true"
		strCountryDisplay = "true"
		strPhoneNbrDisplay = "false"
		strCommentsDisplay = "true"
		strURLDisplay = "false"
		strCompanyReq = "false"
		strStreetReq = "false"
		strAddress2Req = "false"
		strCityReq = "false"
		strStateReq = "false"
		strZipReq = "false"
		strCountryReq = "false"
		strPhoneNbrReq = "false"
		strCommentsReq = "false"
		strURLReq = "false"
		strMbrPswdReq = "false"
		strAcceptCC = "false"
		strDisplayInd = "001110101111010"
		strRequiredInd = "0011100000000000"
	END Select
 
END SUB%>
		
<%FUNCTION CalculateShipping(intNbrItems, intOrderTotal, intTotalShippingWt)
		
	Select CASE SESSION("C")
	  CASE "", "1"
CalculateShipping = 3.50 + (.5 * (intNbrItems - 1))
	End Select
END FUNCTION%>
 
<%SUB Display_Navigation_Graphics
 
	Select CASE SESSION("C")%>
	<%  CASE "", "1"%>
    <div align="center">
      <table border="0" cellpadding="0" cellspacing="0" width="429">
        <tr><!-- Shim row, height 1. -->
          <td><img src="../Graphics/shim.gif" width="47" height="1" border="0"></td>
          <td><img src="../Graphics/shim.gif" width="92" height="1" border="0"></td>
          <td><img src="../Graphics/shim.gif" width="95" height="1" border="0"></td>
          <td><img src="../Graphics/shim.gif" width="109" height="1" border="0"></td>
          <td><img src="../Graphics/shim.gif" width="86" height="1" border="0"></td>
          <td><img src="../Graphics/shim.gif" width="49" height="1" border="0"></td>
          <td><img src="../Graphics/shim.gif" width="51" height="1" border="0"></td>
          <td><img src="../Graphics/shim.gif" width="1" height="1" border="0"></td>
        </tr>
        <tr><!-- row 1 -->
          <td><a href="/index.html" onMouseOut="GrpRestore('FwSimpleGroup');"  onMouseOver="GrpSwap('FwSimpleGroup','nav5_r1_c1','Graphics/nav5_r1_c1_F2.gif')"  onClick="GrpDown('FwSimpleGroup','nav5_r1_c1','Graphics/nav5_r1_c1_F3.gif','Graphics/nav5_r1_c1_F3.gif')" ><img name="nav5_r1_c1" src="../Graphics/nav5_r1_c1.gif" width="47" height="27" border="0"></a></td>
          <td><a href="/bookpg.shtml" onMouseOut="GrpRestore('FwSimpleGroup');"  onMouseOver="GrpSwap('FwSimpleGroup','nav5_r1_c2','Graphics/nav5_r1_c2_F2.gif')"  onClick="GrpDown('FwSimpleGroup','nav5_r1_c2','Graphics/nav5_r1_c2_F3.gif','Graphics/nav5_r1_c2_F3.gif')" ><img name="nav5_r1_c2" src="../Graphics/nav5_r1_c2.gif" width="92" height="27" border="0"></a></td>
          <td><a href="/linenpg.shtml" onMouseOut="GrpRestore('FwSimpleGroup');"  onMouseOver="GrpSwap('FwSimpleGroup','nav5_r1_c3','Graphics/nav5_r1_c3_F2.gif')"  onClick="GrpDown('FwSimpleGroup','nav5_r1_c3','Graphics/nav5_r1_c3_F3.gif','Graphics/nav5_r1_c3_F3.gif')" ><img name="nav5_r1_c3" src="../Graphics/nav5_r1_c3.gif" width="95" height="27" border="0"></a></td>
          <td><a href="/creationpg.shtml" onMouseOut="GrpRestore('FwSimpleGroup');"  onMouseOver="GrpSwap('FwSimpleGroup','nav5_r1_c4','Graphics/nav5_r1_c4_F2.gif')"  onClick="GrpDown('FwSimpleGroup','nav5_r1_c4','Graphics/nav5_r1_c4_F3.gif','Graphics/nav5_r1_c4_F3.gif')" ><img name="nav5_r1_c4" src="../Graphics/nav5_r1_c4.gif" width="109" height="27" border="0"></a></td>
          <td><a href="/wovenpg.shtml" onMouseOut="GrpRestore('FwSimpleGroup');"  onMouseOver="GrpSwap('FwSimpleGroup','nav5_r1_c5','Graphics/nav5_r1_c5_F2.gif')"  onClick="GrpDown('FwSimpleGroup','nav5_r1_c5','Graphics/nav5_r1_c5_F3.gif','Graphics/nav5_r1_c5_F3.gif')" ><img name="nav5_r1_c5" src="../Graphics/nav5_r1_c5.gif" width="86" height="27" border="0"></a></td>
          <td><a href="/searchpg.html" onMouseOut="GrpRestore('FwSimpleGroup');"  onMouseOver="GrpSwap('FwSimpleGroup','nav5_r1_c6','Graphics/nav5_r1_c6_F2.gif')"  onClick="GrpDown('FwSimpleGroup','nav5_r1_c6','Graphics/nav5_r1_c6_F3.gif','Graphics/nav5_r1_c6_F3.gif')" ><img name="nav5_r1_c6" src="../Graphics/nav5_r1_c6.gif" width="49" height="27" border="0"></a></td>
          <td><a href="/linkspg.shtml" onMouseOut="GrpRestore('FwSimpleGroup');"  onMouseOver="GrpSwap('FwSimpleGroup','nav5_r1_c7','Graphics/nav5_r1_c7_F2.gif')"  onClick="GrpDown('FwSimpleGroup','nav5_r1_c7','Graphics/nav5_r1_c7_F3.gif','Graphics/nav5_r1_c7_F3.gif')" ><img name="nav5_r1_c7" src="../Graphics/nav5_r1_c7.gif" width="51" height="27" border="0"></a></td>
          <td><img src="../Graphics/shim.gif" width="1" height="27" border="0"></td>
        </tr>
      </table>
    </div>

	<%End Select
END SUB%>

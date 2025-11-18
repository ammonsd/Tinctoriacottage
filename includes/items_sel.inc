<%
'          **************************************************
'          *  Copyright Fulcrum Computing, Inc.  2000       *
'          **************************************************
'          * This program is copyrighted under the laws of  *
'          * the United States. Any unautorized duplication *
'          * or use of this progrm is prohibited.           *
'          **************************************************
'**Start Encode**#@~^BAAAAA==@#@&FwAAAA==^#~@%>

<!--#INCLUDE FILE="../includes/get_program_path.inc" -->
<!--#INCLUDE FILE="../includes/transfer_html.inc" -->
<!--#INCLUDE FILE="../includes/check_date.inc" -->
<!--#INCLUDE FILE="../includes/checknumeric.inc" -->
<!--#INCLUDE FILE="../includes/get_refdata.inc" -->
<!--#INCLUDE FILE="../includes/build_where_sql.inc" -->
<!--#INCLUDE FILE="../includes/db_setup.inc" -->
<!--#INCLUDE FILE="../includes/company_setup.inc" -->
<!--#INCLUDE FILE="../includes/display_edit_errors.inc" -->
<!--#INCLUDE FILE="../includes/validate_email_format.inc" -->

<%#@~^ZQwAAA==@#@&@#@&,P,fqtPkOD}wOkKxdS,/YMokVNB~dYMATZsDS~o~~qB~/DDPnsw~~dDD(Yh:X2~,/YMnThSKmSPkY.j$V4nD~P(GsqD+s?.m4@#@&@#@&P,~/DDKLsSW^~{P!+DKT:KmY4`Jr#@#@&P,P/l^V~jYEamZK:wmUz{&x6W@#@&~~,mCV^~fmYC8m/+mjY;wvdDD/K:alxH1mh+*@#@&@#@&P~~&0P"n;!+/D }EDH/ODrUT`E?rbP{PEerPW.~"+5EdDRoKDs`JUJ*~',JeJ,PP4x@#@&~P,PP,8GV&Y:jD^4,'~YM;+@#@&~~,2Vdn@#@&~P,~,P8KV&Y+s?M^t,'~0mVdn@#@&P,~2	NP&W@#@&@#@&,P~qW~"+5EdYcp;nMX/O.bxL`rPr#~@!@*,JJ,K4nx@#@&~P,P~~kYD&O+sKXan~',I;;+dOcp;+Mz/DDrUT`JPE*@#@&P,~AVd@#@&PP,P,~/DD(Y:Pza+P{~I;EdORwWM:cJPE*@#@&P,~2	N~(6@#@&@#@&,P~/D.~o/^D,'PrP(LmKVGD{J~',ZC"c&W#PL~EaZ!R!0!E~LP/C"c&W#~',PJ~sbxV'r~LP/_Iv&c*PL~P,J:!Z!TTZJP,'P;CIvf*#,[,PEP-sbxV'r~[,Zu]v&cb~LPEaZTZ!TZJ,P[,Z_]`2cb@#@&@#@&~,Pq6~I;EdORwWM: ZG;	Y~',TP:tnU@#@&P~~,P~q6~"+5!+kYR5E.XUY.k	ocEhIKrbP{PJIE~K4+	@#@&P~~,P~P,~?A?j(}1`EK"KE#,x,Jer@#@&PP,P,~2^/n@#@&P~~,PP,~PU2?U(61vJhIPJb~{JE@#@&~P,P~~AxN~(6@#@&P,~,P~1l^VPGkk2VmXmsKDh@#@&PP,3Vk+@#@&~~P,P,ZCVs~jlsk9CY{(UaEY@#@&,P~P,~&0~DMH/T/c/W!xOP@*PT~:t+	@#@&,PP,~~P,PmDz2..KD\/TdP{Pn.MH/LdcqO+sd@#@&~,P,PP,P,^l^V~fb/2smX{A.DKD/@#@&~P,P,P3x[~&0@#@&,~P,P~dDDr2ObWU/,x,JE@#@&,PP,P,oWMP3l1t~a,kx,]+$E+kO sKDs@#@&P~~,P~P,~q6P]n$E+dOcsGDscX#~{PrJPKD,aP{PEKrPG.,6P{~JUJPK.~6,',JjE8hbYEPK.PXPx~rXJ~GMPaP{~r6E,K4+x@#@&,~P,P~P,P3sk+@#@&~P,PP,~~P,P,/ODorV[P{~6@#@&~~,PP~~,P~P,(6Pa,',J1KYz2wMW-+9J~P4+x@#@&P,PP,~~P,P,P~P~~kY.sbnV9Px~rbw2.K\CVGCD+E@#@&,PP,P,~P,P~P,P~~kYD:n:aP',E J@#@&,P~P~~,P~P,~2^/nr6PVnWD`a~8+*Px,JzwwMW7CVGlO+rPP4x@#@&~P,PP,~~P,P,P~P(W,In;!n/DRoGM:`EHKY)wa.K\n9J*P',Jr~K4+U@#@&P~~,PP,~P,PP,~~P,P&0~kdH!:nDb^`"+5;/Y oKDh`Xb*Px,YME+,Y4nx@#@&~P,P~~,PP,~P,PP,~~P,PkY.KnhaPxP;G:aEOnGlYnc"+5EdDRoKDs`6*#@#@&P,P~P,P~~,PP,~P,PPAsd+@#@&,P~P~~,P~P,~P,P~~,PP~dDDP+s2,'~;tm3GlDn`"+5E/O wWDsc6*#@#@&~~P,P,P~P~~,P~P,~Px[~&0@#@&~,P~P,~,P~,P,PPAVkn@#@&P~P,P~~,PP,~P,PP,~dYMsb+sN~x,JE@#@&~P,P~~,PP~~,P~PAU9P(6@#@&PP,P,~P,P~P,2sdq0,aP{PJ&On:&frPGD~a,'~J5;l	YrOHJPG.,6~',EhDr1+rPK4+	@#@&,P~P,P~~,PP,~P,/YMPn:aP{PODrhvInw^/tmDdc"+;;nkY sK.s`a*~r~JBJ,E#*@#@&P,P~~,PP,~P,PP,(WPXP{PEqOnsq9J,Ptx@#@&,PP~~,P~P,~,P~,P,P/DD:n:aPxP!mCd`I2V;tlMdc/DD:+hwSE,JSJBE#*@#@&~,PP~~,P~P,~,P~Ax9Pq6@#@&~P,P~P,P~~,PP,~/DDsbnsN,',J(1 E,[~/D.sb+s[@#@&P~~,P~P,~,P~AVk+q6PX~',JnOtlrsrPK4nx@#@&P,~~P,P,P~P~~,PdYMP+sw~x,I+2s;tCDkcdZCk+vI+$EdYcsGDs`ab*~J,E~r_J*@#@&P,P,P~P~~,P~2^d+&0~a,'PEjl.m4E,K4x@#@&P,P,~P,P~P,P~~,/YMokVN,x~JUJ@#@&~P~~,P~P,~P,P~~kYDPnsw~',]ws;tmD/vKMr:vIn;!+dOcsWMh`X##BE~JBJ3Jb@#@&~,P~P,~P,P~~,PPdOMKn:a~{P]w^ZtmDkc/DDP+swSEuJ~rMJ*@#@&,~~P,P,P~P~~AVd+@#@&P,P~~,PP~~,P~P,dDDP:aP',I2V;tCDk`P.b:`"n;!+/D oWM:v6b#SE,JSJ3E#@#@&~~,PP~~,P~P,3	N~&0@#@&P,P,~P,P~P,P(W,/YMokVN,@!@*PrJ,K4+U@#@&P~P,~P,P~~,PP~~kY.raObWUkP{P/DD}2YbWU/,[~ELJPL~IwV;4CDk`kY.srn^NSJ,E~r_Eb,[PExrP'PkOMKnsw@#@&P,P,~P,P~P,P3U9Pq6@#@&,PP,~~P,PAx[P(W@#@&~P,~P,1naD@#@&~~,P~P&W,4G^qD+:UD14P{POD!+~C	NP"n;!+/D oWM:vJj+C.1tE#{~JrPP4x@#@&~,P~P,~,PdDD}wYbW	dP{PdYMr2ObWxk~[,J[r~'Pr?{RE@#@&~,P~P,3x9P(W@#@&P~~,P~q6~(Ws&Y:?Mm4~',Y.EPP4x@#@&~P,PP,~~PkYMr2YrG	/~',EgrP'~skNcdDD6wDrKxdB *@#@&,P,~P,2s/@#@&~,PP,~P,P/D.6wDkKxdPx~rgP'r~[,/O.&Y+hPHwnPL~kY.}wDkW	/@#@&P,P~P,2U[,q0@#@&@#@&PP,~~P"+kwGxdncInNb.+1Y~E4YY2lJzEPL~kY.hosSW1PL~J^kdY|kOns/RmdwrPPL~dYMraYrWUd@#@&~P,3x9P(W@#@&igMDAA==^#~@%>

<%#@~^EAAAAA==j`A~GkkwVmX|oWM:0wUAAA==^#~@%>

<html>

<head>
<title><%=#@~^DgAAAA==dDD/K:alxH1mh+sQUAAA==^#~@%> Item Selection Menu</title>
</head>

<body <%=#@~^CQAAAA==dDD$TZ^DPQwMAAA==^#~@%> >

<div align="CENTER">
<table border="1" cellspacing="0" cellpadding="2" bordercolor="#000080">
<tr>
<td align="CENTER" valign="TOP" bgcolor="#000095">
<font face="Arial" size="3" color="#FFFFFF">
<%#@~^GgAAAA==(6P8KV&Y+s?M^t,'~YMEn~:t+	FwkAAA==^#~@%>
&nbsp;
<br>
<%#@~^BAAAAA==3^/niQEAAA==^#~@%>
<b>Criteria for Selection</b><br></font>
<font face="Arial" size="2" color="#FFFFFF">
<b>Leave Blank for All Entries</b>
<%#@~^BgAAAA==3	N~&05gEAAA==^#~@%>
</font>
</td></tr>

<tr>
<td valign="TOP" bgcolor="#B4B4B4">
<font face="Arial" size="2">

<div align="CENTER">
<table border="0" cellspacing="0" cellpadding="5">
<form action="http://<%=#@~^CQAAAA==dDDKT:dWmmwMAAA==^#~@%>items_sel.asp" method="POST">
<input type="Hidden" name="T" value="<%=#@~^CwAAAA==dDD(D+sKXa+igQAAA==^#~@%>">
<%#@~^GgAAAA==(6P8KV&Y+s?M^t,'~YMEn~:t+	FwkAAA==^#~@%>
<input type="Hidden" name="S" value="Y">
<%#@~^BgAAAA==3	N~&05gEAAA==^#~@%>

<font face="Arial" size="3">

<%#@~^GgAAAA==(6P8KV&Y+s?M^t,'~YMEn~:t+	FwkAAA==^#~@%>
  <tr><td>
  <font size="-1">
  <b>Keyword Search</b>
  </font>
  </td>
  <td nowrap>
  <input type="Text" name="Search" size="55">
  </td></tr>
  <tr><td>
  &nbsp;
  </td></tr>
<%#@~^BAAAAA==3^/niQEAAA==^#~@%>
  <tr><td nowrap colspan="2">
  <%#@~^GgAAAA==(6PdDD&Y+sKH2+,'~JhJ~P4+x,OQgAAA==^#~@%>
  <b>e-Mail</b>
  <input type="Text" name="e-Mail" size="40">
  <%#@~^BgAAAA==3	N~&05gEAAA==^#~@%>
  <b>Category</b>
  <%#@~^/wAAAA==(6PdDD&Y+sKH2+,'~JhJ~P4+x@#@&P,PP,/CV^PV+OInWGlOlvEZzK3!}I5ESrn6?:E*@#@&,PAV/q6~/DD(Y:Pza+P{~J&JP:4nx@#@&,P~P~/mVsPVnY"+W9mYlcE;bP2V6"5EBJ&1.A1:6IIJb@#@&P~3	NP&W@#@&PP&~xPmDHIn09CDlc!*~_,F@#@&,Pl.z"+WfmOm`(*P{PJgWD~bk/ro	+[E@#@&P,CDHI+69CYm`Z#~'~(@#@&~PP0EAAA==^#~@%>
  <select name="Category" size="1">
  <option selected value=""></option>
  <%#@~^GgAAAA==oKD~&P{PF,YK~lMX]+6fCOm`!*xQcAAA==^#~@%>
  <option value="<%=#@~^DQAAAA==CMX]0GlYm`&bfQQAAA==^#~@%>"><%=#@~^DQAAAA==CMX]0GlYm`&bfQQAAA==^#~@%></option>
  <%#@~^BAAAAA==H6OnwEAAA==^#~@%>
  </select>
  <%#@~^GQAAAA==(6PdDD&Y+sKH2+,'~JhJ~P4+xGQgAAA==^#~@%>
  </td></tr>
  <tr><td nowrap>
  <font size="-1">
  <b>From Approved Date</b>
  </font>
  </td>
  <td nowrap>
  <input type="Text" name="ApprovalDate1" size="10">
  <font size="-1">
  <b>To Approved Date</b>
  </font>
  <input type="Text" name="ApprovalDate2" size="10">
  <font size="-1">
  <b>Not Approved</b>
  </font>
  <input type="checkbox" name="NotApproved" value="Yes">
  </td></tr>
  <tr><td nowrap>
  <%#@~^BgAAAA==3	N~&05gEAAA==^#~@%>
  <b>Web Page</b>
  <%#@~^GQAAAA==(6PdDD&Y+sKH2+,'~JhJ~P4+xGQgAAA==^#~@%>
     </td><td nowrap>
   <%#@~^1wAAAA==/mVs,MYI0GCYm`EAAK!r~Jh6?:J#@#@&~PAVk+(0~dDD(YhKHwn~{PJ(E,K4+	@#@&P~,P,Zl^V,!+DIn0GlOCvJA$nVJ~r(H.A1:r]5Eb@#@&~PAUN,qW@#@&PP(~{PCDH]09mYm`!*P3~F@#@&~PmDz]0fmOlvq#,x~JgWDP)/drTxnNr@#@&,PC.HI+W9mYC`Zb,'~&@#@&PPnTYAAA==^#~@%>
  <select name="WebPg" size="1">
  <option selected value=""></option>
  <%#@~^GgAAAA==oKD~&P{PF,YK~lMX]+6fCOm`!*xQcAAA==^#~@%>
  <option value="<%=#@~^DQAAAA==CMX]0GlYm`&bfQQAAA==^#~@%>"><%=#@~^DQAAAA==CMX]0GlYm`&bfQQAAA==^#~@%></option>
  <%#@~^BAAAAA==H6OnwEAAA==^#~@%>
  </select>
  <%#@~^GQAAAA==(6PdDD&Y+sKH2+,'~J&J~P4+xEggAAA==^#~@%>
  <b>Include Inactive</b>
  <input type="checkbox" name="Inactive" value="Yes">
  <%#@~^BgAAAA==3	N~&05gEAAA==^#~@%>
  </td></tr>
  <tr><td nowrap>
  <%#@~^GQAAAA==(6PdDD&Y+sKH2+,'~JhJ~P4+xGQgAAA==^#~@%>
  <b>Item ID(s)</b>
  <%#@~^HQAAAA==3^/n&0,/YMqDn::X2+,'~E&JP:4+	mwkAAA==^#~@%>
  <font size="-1">
  <b>Product Code(s)</b>
  </font>
  <%#@~^BgAAAA==3	N~&05gEAAA==^#~@%>
  </td>
  <td nowrap>
  <%#@~^GQAAAA==(6PdDD&Y+sKH2+,'~JhJ~P4+xGQgAAA==^#~@%>
  <input type="Text" name="ItemID" size="55">
  <%#@~^HQAAAA==3^/n&0,/YMqDn::X2+,'~E&JP:4+	mwkAAA==^#~@%>
  <input type="Text" name="ProductCode" size="55">
  <%#@~^BgAAAA==3	N~&05gEAAA==^#~@%>
  </td></tr>
  <tr><td nowrap>
  <b>Title</b>
  </td>
  <td nowrap>
  <input type="Text" name="Title" size="55">
  </td></tr>
  <tr><td nowrap>
  <b>Description</b>
  </td>
  <td nowrap>
  <input type="Text" name="Description" size="55">
  </td></tr>
  <tr><td nowrap>
  <b>Keywords</b>
  </td>
  <td nowrap>
  <input type="Text" name="Keywords" size="55">
  </td></tr>
<%#@~^BgAAAA==3	N~&05gEAAA==^#~@%>
<tr>
<tr><td colspan="2" nowrap>
<hr>
</td></tr>
<td colspan="2" align="CENTER" nowrap>
<input type="Submit" value="Submit Request">
&nbsp;&nbsp;
<input type="Reset">
&nbsp;&nbsp;
<a href="http://<%=#@~^EAAAAA==dDD/K:alxHCKh+hocAYAAA==^#~@%>/registry_doc.html#Items" target="Help_Win"><img src="../graphics/buttons/help.gif" width="12" height="16" border="0" align="top" alt="Visitor Registry Documentation"></a>
</td></tr>
</font>
</form>
</table>
</div>
</td></tr>
</table>
</div>
</body>
</html>

<%#@~^CAAAAA==3	N~UE(PgQIAAA==^#~@%>

<%#@~^MgUAAA==j`A~jl^kNmYmq	w;Y@#@&@#@&,PBJ&PjlVb[CYP6W.:~n	Y.kdR@#@&@#@&,PP9(tP8W^3MDS,k	Y2MDg8D@#@&@#@&,P~jYP.Dt/ok~xP;DlO+68N+^YvE?1Dr2DkxL Gk^YbG	l.HJ*@#@&@#@&,~P&0~YMkhc"+;!n/DRsK.h`rbaw.W-C^fCYqJ*#~@!@*PJE~:tnx@#@&,P~,P,/YMKhw,'~YMkhc"+;!n/DRsK.h`rbaw.W-C^fCYqJ*#@#@&,PP~~,qWPbd9lO`kYD:+s2#,'~0mVdn,YtU@#@&PP,~~P,P(Ws2..,'~YM;+@#@&~~,PP~3^/nq6~wW.slDflD+:r:`dYMKnha#P@!~sKD:mO9lD+:kh+cEZF&!8&FO!TE*PY4n	@#@&P,~,P~,P(WVADM~',Y.E@#@&~,PP,~2	NP&W@#@&,P,P~P(W,4GVA.D,'~OME+~O4+U@#@&~,P~,P,PPbxD3DM18D,'~r	Y2M.1(DP3~q@#@&P,P~P~~,PnDM\/T/ )9NPrUD2.Dg8M~~rq	\l^k9~sMWhPzw2.K\l^~fmY+r@#@&P,P,P~2U[,qW@#@&~P,2U[,q0@#@&,P~q6~DDrs`"+;!+kORwW.:vJ)2aDW7CVGlY+E#*P@!@*~JE~:tnx@#@&P,P~~,/Y.P:2P{~DDrs`"+;!+kORwW.:vJ)2aDW7CVGlY+E#*@#@&P~P~~,qWPbdNmYnckYDPnswbP{~6lsk+,Ytx@#@&P,P~P,P~8KV2M.P{PYM;n@#@&P,P~P~3^/nq6~sKDhCDflOn:kh+vdDDP:a#P@!PwGDslOfmYnPb:+vE!8z!8&q,Z!r#~Y4n	@#@&P,~P,P~~(WV3.MPxPD.!+@#@&P,PP,PAUN,qW@#@&P~~,PP&WP(WVA..P{PDD;+~O4+U@#@&~P,P~~,PPrUD2.Dg8MPx,k	Y2MDg8D,_~F@#@&~~,PP,~P,+DM\dokRzN[PrUD2.Dg8DBPE(	\lsr9PPW,)aw.K\mVPGlDnJ@#@&~P,P~~AxN,(0@#@&P,~3x9P&0@#@&~~,qWPD.ks`]n$E+dOcsGDscrqO:&fJ*#,@!@*,JEP:tnU@#@&P,~P,Pq6~/tm01;:n.bmcI5E/O wWDhcrqO+s(GJbBJBPOr#,xP6ls/PP4x@#@&~P,PP,~~PbxD2.DH8MPxPbUYAD.H(DPQ~8@#@&P,~,P~,P,+DMHkL/cb[N,kUOADDg8DBPJ&On:,qG/~:;dDP8+,UEs+.r1J@#@&~,P~P,3	N~&0@#@&P,PAUN,qW@#@&@#@&3	NPUiAI2IBAA==^#~@%>


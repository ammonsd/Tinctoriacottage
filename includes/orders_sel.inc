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
<!--#INCLUDE FILE="../includes/company_setup.inc" -->
<!--#INCLUDE FILE="../includes/state_country_select.inc" -->
<!--#INCLUDE FILE="../includes/display_edit_errors.inc" -->
<!--#INCLUDE FILE="../includes/validate_email_format.inc" -->
<!--#INCLUDE FILE="../includes/checknumeric.inc" -->
<!--#INCLUDE FILE="../includes/replace_chars.inc" -->
<!--#INCLUDE FILE="../includes/check_date.inc" -->

<%#@~^TAgAAA==@#@&@#@&,P,fqtPkOD}wOkKxdS,/YMokVNB~dYMATZsDS~kY.nThSKmS~o~P(S,/OD:nsw@#@&@#@&PP,/D.nT:JW1Px~V+YhL:hlY4cEJ*@#@&P~P/C^V~?OEa{/GswlUz|qU0K@#@&@#@&,P,?2U?&61vJ\/TJb~{PJr@#@&@#@&P,~dYMATZsD~x,J~4T^W^W.xrP[~/_Ic&Wb,[~raZ!%Z%ZEPLP/C"`f**P[,~J,Vk	VxJ,[,ZuIcfW#~[,~PraTTZ!!TE,P'P;u"`fW#,[P,J,-VbxV'rP'~;CIvfc*P[,E:!Z!Z!TJ~~LP/C"c&W#@#@&@#@&P~~&0~I5!+dDRwWDsR;GE	Y~',!~P4+x@#@&P,PP,~(0,I;;+dOcp;+Mz?DDrUT`JK]:JbP{~r5E,K4+x@#@&,~P,P~P,Pj3U?q}H`rnI:EbP{Pr5E@#@&~,P~P,3Vk+@#@&,PP~~,P~PU3U?(}1vJn"KrbP{JE@#@&P~~,PPAUN,q0@#@&~P,P,P^lss,fr/aslH{oGM:@#@&~,P3Vkn@#@&~,P,PP;l^sPjlsk9lOn|qxa;Y@#@&P,~~P,q6PnD.\kodR;GE	Y~@*,!PP4x@#@&,~,P~,P,PlMXA.DKD\/T/~x,+DM\/T/R&On:k@#@&P~P~~,P~P1CV^P9rkwVCz|2.DK.k@#@&,P,PP,2	[P&0@#@&,P~~,PsK.PAlm4~aPbx,In;;nkY sK.:@#@&~~,PP~~,P(0,cXPx,JUrJ,WM~6,'~J;rEb,lx9~I;EdORwWM:c6b~{PEqr~K4+U@#@&PP~~,P~P,~,PvGW,1WDPh.W1+d/,`[n6lE^OP7lV!nb@#@&P,P~P~~,P3Vknq6P]n$E+dOcsGDscX#~@!@*,JJ,l	[PXP@!@*,JHnXYnThJ,KtU@#@&,P,P~P~~,P~P&WP^+WOv6~b,'~JRtlr^J,Ktx@#@&P,P~P,P~~,PP,~PkYD:nhw,',Inws/4l./vJZm/nc"+;;nkY sK.s`a*#BJPr~rQJ*@#@&P,P~~,PP,~P,2Vkn(0,k	/ODc;1ld+va#BJ9):2Jb~@*PTP:4x@#@&P,PP,P,~P,P~P,PdOMK+s2P{PZ4n^3GlD+cIn5!+dYcoWM:ca*#@#@&~,P~P,~,P~,PAV/@#@&~P,P~P,P~~,PP,~/DDKh2P{P"+2V/4mDd`"n;!+dOcsW.hv6b~r~r~E3J*@#@&,P,~P,P~P,P~3	NP&W@#@&PP,~~P,P,P~PdOMsr+^[P{Pa@#@&PP~~,P~P,~,P(6PkYDwksN,'~J\k2E,KtU@#@&PP,~~P,P,P~P~~,/OD:n:aPx~^+0OckY.Kha~X*@#@&PP,P,~P,P~P,P~~&0PbU/DD`kO.K:a~EeEb,'~!,Ptx@#@&,PP~~,P~P,~,P~,P,P/DD:n:aPxPkY.P:w,'PreJ@#@&~P,P,P~P~~,P~P,3x9P(W@#@&P~~,P~P,~,P~AVk+q6PkODwknV9Px~rrD9nD[JP:4nx@#@&,P~P~~,P~P,~P,PdOMskns9PxPr6gJ@#@&P,PP,P,~P,P~2	N~(6@#@&,~P,PP,~~P,P&0~/O.}wOkKU/,'~ErPK4n	@#@&P,~,P~,P,PP,P,~/DD6wDkGUkP',Egr@#@&,~~P,P,P~P~~AVd+@#@&P,P~~,PP~~,P~P,dDD6aYbWxkP{~/DD6wDkGUkP[,E[r@#@&,~~P,P,P~P~~Ax[P&W@#@&P~~,PP~~,P~PkOMr2DkKx/,',dYMr2YbWUd,[P"nw^Ztm.d`kYMsr+s[BJ~JBE_r#~',J'E~LPdYMP:2@#@&,PP,P,~P,2UN,qW@#@&PP,~P,1+XO@#@&,P,P~P(W,In;!n/DRoGM:`EH6OnThr#~{Pr2J,K4nx@#@&~P,P~~,PPkOD:+:a~xPrWMNnDdmkYCY!d{:Cr^Rld2r@#@&P,~,P~AVk+@#@&P,~P,P~P,/O.:+:a~',JVbdO{KD9+./ CkwE@#@&~P,P~~AxN~(6@#@&P,~,P~"+kwW	/ INrDmO~rtYD2)JzJ,'~/DDhohSG^,[~/D.K:2~,[PdOMr2YbG	/@#@&P,P2	N,(0@#@&8PgBAA==^#~@%>

<%#@~^EQAAAA==j`A~GkkwVmX|oWM:~8wUAAA==^#~@%>

<html>

<head>
<title><%=#@~^DgAAAA==dDD/K:alxH1mh+sQUAAA==^#~@%> Orders Selection Menu</title>
</head>

<body <%=#@~^CQAAAA==dDD$TZ^DPQwMAAA==^#~@%> >

<div align="CENTER">
<table width="575" border="1" cellspacing="0" cellpadding="2" bordercolor="#000080">
<tr>
<td width="50%" align="CENTER" valign="TOP" bgcolor="#000095">
<font face="Arial" size="3" color="#FFFFFF">
<b>Criteria for Selection</b><br></font>
<font face="Arial" size="2" color="#FFFFFF">
<b>Leave Blank for All Entries</b></font>
</td></tr>

<tr>
<td valign="TOP" bgcolor="#B4B4B4">
<font face="Arial" size="2">

<div align="CENTER">
<table width="90" border="0" cellspacing="0" cellpadding="5">
<form action="http://<%=#@~^CQAAAA==dDDKT:dWmmwMAAA==^#~@%>orders_sel.asp" method="POST">
<input type="Hidden" name="NextPgm" value="<%=#@~^IAAAAA==i;ld`"+;!+kOR5EnDH?O.bxovEnr##,ywoAAA==^#~@%>">
<font face="Arial" size="3">

<tr><td>
<b>Customer Name</b>
</td>
<td nowrap>
<input type="Text" name="Name" size="50">
</td></tr>

<tr><td nowrap>
<font size="-1">
<b>From Shipped Date</b>
</font>
</td>
<td nowrap>
<input type="Text" name="ShippedDate1" size="10">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font size="-1">
<b>To Shipped Date</b>
</font>
<input type="Text" name="ShippedDate2" size="10">
<font size="-1">
</td></tr>

<tr><td nowrap>
<font size="-1">
<b>From Order Date</b>
</font>
</td>
<td nowrap>
<input type="Text" name="OrderDate1" size="10">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font size="-1">
<b>To Order Date</b>
</font>
&nbsp;&nbsp;&nbsp;
<input type="Text" name="OrderDate2" size="10">
<font size="-1">
</td></tr>

<tr><td nowrap>
<b>State</b>
</td><td nowrap>
<%#@~^GgAAAA==/mVs,?DlY{UnVmOkKxmJb/Y,vQkAAA==^#~@%>
<b>Zip Code</b>
<input type="Text" name="Zip" size="10">
<b>Country</b>
<%#@~^HAAAAA==/mVs,ZKExDDHm?VnmDkGU|SkkOPsAoAAA==^#~@%>
</td></tr>

<tr><td nowrap>
<b>Order Number(s)</b>
</td>
<td nowrap>
<input type="Text" name="Order#" size="50">
</td></tr>

<tr><td nowrap colspan="2">
<b>Pending Shipments</b>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<b>Include?</b>
<input type="Radio" name="SO" value="I" checked>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<b>Exclude?</b>
<input type="Radio" name="SO" value="X">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<b>Only?</b>
<input type="Radio" name="SO" value="O">
</td>
</tr>

<tr><td nowrap colspan="2">
<b>Canceled Orders</b>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<b>Include?</b>
<input type="Radio" name="CO" value="I" checked>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<b>Exclude?</b>
<input type="Radio" name="CO" value="X">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<b>Only?</b>
<input type="Radio" name="CO" value="O">
</td>
</tr>

<tr>
<tr><td colspan="2" nowrap>
<hr>
</td></tr>
<td colspan="2" align="CENTER" nowrap>
<input type="Submit" value="Submit Request">
&nbsp;&nbsp;
<input type="Reset">
&nbsp;&nbsp;
<a href="http://<%=#@~^EAAAAA==dDD/K:alxHCKh+hocAYAAA==^#~@%>/registry_doc.html#OrdersSel" target="Help_Win"><img src="../graphics/buttons/help.gif" width="12" height="16" border="0" align="top" alt="Visitor Registry Documentation"></a>
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

<%#@~^CAsAAA==j`A~jl^kNmYmq	w;Y@#@&@#@&,PBJ&PjlVb[CYP6W.:~n	Y.kdR@#@&@#@&,PP9(tP8W^3MDS,k	Y2MDg8D@#@&@#@&,P~jYP.Dt/ok~xP;DlO+68N+^YvE?1Dr2DkxL Gk^YbG	l.HJ*@#@&@#@&,~P&0~YMkhc"+;!n/DRsK.h`r?4k2wn[GlO+8E#*P@!@*,JJ~P4+U@#@&~,P~,PkYD:+s2P{PODb:c];EdYcsWMhcJUtbw2+[9mYnFrb#@#@&~~,PP~(6Pr/9CD+ckYMK+sw*~',0CVk+~O4+x@#@&P,PP,~~P(W^2.D~x,Y.E@#@&,P~~,P2sdqWPwGM:CDfmY+:ksn`kY.K:2b,@!PwGDslYGCO+:ks+cJTqJ!qz81!ZJb~Dt+U@#@&P~P,~,P~(W^2DMP{~YMEn@#@&P~~,PPAUN,q0@#@&~P,P,P(0~8KV3DM~',Y.;PY4n	@#@&P,~,P~,P,kxD2M.1(D~',kUOADDg8D,_P8@#@&P,P,P~P~~D.HkL/cb[[,kxO3MDH4MS,J(	\mVk9Pw.WsPjtbw2n9PfmO+r@#@&,~~P,PAx[P(W@#@&~P,3x9P(W@#@&P~~&0~YMrs`];!+/DRwGDs`E?4k22NfmO+yJ#*~@!@*,JrPPtnU@#@&~P,~P,/O.:+:2~{PODbhvIn$E/YcsK.:vJjtbw2n9flDn r##@#@&~P,P,P(0~rkNCYc/DDPnsw#~x,0CVkn,Y4x@#@&P,P,~P,P8W^2..,'PD.E@#@&,~~P,PAVd+(W,sGDsCYGlOn:k:nckY.Kha#~@!PwWDslD9lD+Pks+cEZFzZqz8,!ZEbPDtx@#@&~~,P~P,~4KV3.MP'~OMEn@#@&~,P~,PAxN,q6@#@&,P~P,P(W,4W^3DMP',O.EPDtnx@#@&,P~P,~P,PrUD2D.H(D~',r	Y3MDg4D,_,q@#@&P~P,P~~,P+M.Hko/c)[N,k	Y3D.H(DSPr(x7lsr9PKG~Utrwan9P9mYJ@#@&P,~P,P3x9P(W@#@&P,~2	NP&W@#@&@#@&,P~qW~DDr:v]+$EndDRsG.s`ErM[D9mYFJ*#,@!@*,JEP:tnU@#@&P,~P,P/D.P+sw,'~Y.rs`]+$;+kY oKD:cE}D[+M9mYn8J*#@#@&P,~P,P(0,kd[mY+vdYMK+s2bP{P6ls/n~Dtnx@#@&P,P~~,PP8G^2.D,x,Y.!+@#@&P,P,~PAVd+&0~oKD:mOfmY+:rh+v/DDP+h2*P@!PwGDslO9mY+Prs+cJZqJ!qJFO!!r#,Otx@#@&,P~~,PP,8W^2DM~xPDD!+@#@&~~,P~PAUN,qW@#@&PP~~,P(0,8KV3MD,'PDD!nPDtnx@#@&~~,PP,~P,kxD3.Dg4MPxPrUD2.Dg8D,_~q@#@&P~~,P~P,~D.t/T/RzN9~k	Y3DM18.BPJ&U\mVk9~oDK:,r.Nn.,fCYE@#@&P~~,PP3U9P(0@#@&,P~Ax9Pq6@#@&~P,qWPDDrhvI+$;+kYRwG.:vJ}D[+.9mYn rb#,@!@*~rJPP4x@#@&,~,P~,/DDK:a~',Y.ks`]n$E+kORwWDscErMND9lOnyJb#@#@&P,P~~,q0~rkNCYckY.:+sw#,',Wl^/nPDtnU@#@&P,~P,PP,8GVADMPxPO.!+@#@&,~P,P~3^/+(W,sGDsCDfCD+:k:`kOD:+hw*P@!~wWDsCYGlYPr:`r!qzTqJF1!ZE#,Y4n	@#@&~~,P~P,~(WsADMP',YM;+@#@&~P,P~~AxN,(0@#@&P,~~P,q6P8Ws3MD~',OD!+~O4+x@#@&,P~P,~,P~bxD2DM1(.P{PrxD2..g4D,QP8@#@&,~~P,P,P~+..t/L/c)N9PrUD2D.H(DSPr(	\C^k9PKKP}.ND~fmYnE@#@&P,~P,P2	[~q6@#@&P~P3U9P(0@#@&@#@&P~~&0PO.b:cI5!+dDRwWDs`rtkaJb#,@!@*~rJP:4+	@#@&,~~P,P&0~VnUvIn;!n/DRoGM:`EtbwE#*~@!PX,WMPVxv]+$En/DRoGM:`rtkaJ#*~@*P8!,W.P/4mV1!h+Mk^c"+;;nkY sK.s`E\kaJ#BJ E#,'~0mVdn,KtU@#@&PP,~~P,P,kUY3.M18D,xPbxO3MD18.,_~F@#@&,P~,P,PP,+M.HkodRzN[~bxYA.Dg4DB~Eq	\mVrN~tbw~ZK[+,sG.slYE@#@&P~P,~,2sk+&0P^+	cI;;+kY oKD:vE}bwJ*b~@*,*,K4+U@#@&P~P,~P,P~(6P:r[vIn;!nkY wWM:`r}b2J*~~8#~@!@*PJ EPKDP^nU`"+$En/O wW.:vE}bwEb*P@!@*~8!~K4n	@#@&,P,PP,P,~P,PrxD2..g4D,xPbxYA..1(D,_~F@#@&,P~P,~P,P~~,+D.\kodRz[9Pr	YADDg4MSPrqU\mVr[,}kaQc,sWMhCY,`	xUxUR	xUx*E@#@&P~~,PP~~,2UN,(6@#@&,P,PP,2	[P&0@#@&,P~3	NP&W@#@&@#@&,~~q6PDDr:c];;+kORwW.hvJr.[D:J*b,@!@*,JrPK4+	@#@&,P~P,P(W,Zt^3gE:.rmvI;;+dOcsGDscJ}D[nMaJbSr~~Orb,'~6l^/+,K4nx@#@&~P,P~~,PPbUYADDg8.P{PbxO2..g4.P3~F@#@&~~,PP~~,PnDM\kodcb9NPbxD3DM18DBPE6MN+M~1!:4.dPsEkY~4n~	Eh+Mrmr@#@&~,PP~~Ax[P&W@#@&~,PAxN,q6@#@&@#@&3x9Pji~9egCAA==^#~@%>

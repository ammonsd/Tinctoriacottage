<%
'          **************************************************
'          *  Copyright Fulcrum Computing, Inc.  2000       *
'          **************************************************
'          * This program is copyrighted under the laws of  *
'          * the United States. Any unautorized duplication *
'          * or use of this progrm is prohibited.           *
'          **************************************************
'**Start Encode**#@~^BAAAAA==@#@&FwAAAA==^#~@%>

<!--#INCLUDE FILE="../includes/db_setup.inc" -->
<!--#INCLUDE FILE="../includes/items_tbl_setup.inc" -->
<!--#INCLUDE FILE="../includes/build_where_sql.inc" -->
<!--#INCLUDE FILE="../includes/logon_check.inc" -->
<!--#INCLUDE FILE="../includes/get_program_path.inc" -->
<!--#INCLUDE FILE="../includes/get_words.inc" -->
<!--#INCLUDE FILE="../includes/build_data_array.inc" -->
<!--#INCLUDE FILE="../includes/company_setup.inc" -->

<%#@~^qAMAAA==@#@&@#@&,P,Pf&H,dYM?5Vqtn.~PoSP&~PkO.K:a~~q(S,/ODGCYm~~6(LsUS,W8L:nXYo	~,/YMKH2+hDGm/dS,/YMPXa+SbdO~,/DDKWdON(fk@#@&@#@&ji~PA;r^Nmq	-xOKDH{Sb/DcKHwnSb/OSq+4hL#@#@&@#@&~~P,/DDKohJKm~',!+DnLhhlY4crJb@#@&@#@&P~,P;lV^PUnY!wmZK:2C	X{&U0K@#@&,~~P1l^V~fCOm4C/m?Y;2v/Y./K:2l	zglh#@#@&@#@&P,~P1lsV,SGLKx{;4+13`r8;k^N|kU\nUDW.X|A+(wL m/wESkY.ZKhalUH1m:+*@#@&@#@&,P~PkY.PHw+dr/DP',PzwSb/O@#@&~,P~/D.4KL,'PJ/m/n`qn(n!*@#@&@#@&,P,~q6P]+$EndDRsK.:cZW!UOP{PZPPtnU@#@&~P,~P,P^C^VP9rkwslHm;tC	o/@#@&P,~PAVd+@#@&~~,PP,~q6PI5;+kYcsGDhcrqU\UHmkUOr#P@!@*,JEP:4x@#@&P,PP,P,~P,In/aWUdRI[kM+mDcEtDYa)&zE~LPdYMKosSG^,[PErD+h/|dV m/agKHwxqr#@#@&,P~~,PPAs/@#@&,~~P,P,P~PdOMKzwKDKmndkP'~EwJ@#@&,~,P~,P,PP1l^sP~ErV9{(H;{sbs+@#@&P,~~P,P,P~mCs^P$EbsN|qU-xYG.H{)DMCH@#@&,P,PP,P,~P"+dwKxdncI+9rDmYvE4YDw=z&J~',In;!n/DRjnM\+.#mDrl(s/cr?AI.AI|Hbt2E#,[~EJJPL~/DD8Ko,[,J /4OsVE#@#@&P,P~~,P2U[,qW@#@&~,P~Ax9Pq6@#@&31GPjj~@#@&@#@&?j~~fb/w^Cz{;tmxL+dmAMBAA==^#~@%>

<html>
<head>
<%#@~^IAAAAA==(6PjA?Uqrg`r$o;V.J*P@!@*,JJ,Ptx/wcAAA==^#~@%>
<body bgcolor="<%=#@~^EAAAAA==jA?j&rg`J~o;sDr#gwQAAA==^#~@%>">
<%#@~^BAAAAA==3^/niQEAAA==^#~@%>
<body>
<%#@~^BgAAAA==3	N~&05gEAAA==^#~@%>
<title>Web Page Changes</title>
</head>
<h4>
After reviewing the proposed changes,<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&#149;&nbsp;
APPLY the changes to the Web Page, or<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&#149;&nbsp;
RETURN to the Maintenance Window
</h4>

<form Action=build_inventory_webpg.asp?T=<%=#@~^CwAAAA==dDDPHwSkkYlwQAAA==^#~@%>&WP=<%=#@~^CAAAAA==dDD	4hoLgMAAA==^#~@%>  Method=POST>
<%#@~^EwAAAA==^mVs,A!kV9{&HZ|srVxAYAAA==^#~@%>
<br><br>
<center>
<input type="Submit" name="Apply" value="Apply Changes">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="Submit" name="InvenMaint" value="Inventory Maintenance">
</form>
</center>
</body>
</html>

<%#@~^CAAAAA==3gf~Uj~PAQIAAA==^#~@%>

<%#@~^ugsAAA==j`A~~EbVN|qg/{wks+@#@&@#@&,PP,(0,/YMPzwnMW^+dd,'~JwEP:tnU@#@&P~~,P~PkOMfCDl,'PU+M-+MR\lawCO4`J'rx1VE9nd-r#,[~JwE,[~V1C/`dOM+8KT#~[,Ec/dbJ@#@&P,P,~P,?nY,W8%wxP{~ZM+lDn64N+1YcJj^Mk2YbUocsrs?XdO:64Nn1YE*@#@&PP,P,~PU+OP}4%P6YwUP{Pr(%oxcZM+CYnP6Osbs+v/O.GlYCS,Y.Eb@#@&~,P,2x9P&W@#@&@#@&P,P~jYPK8L"+mK.[?Y,'~?n.7+.R;.+mYn6(L+^OvJ)f}9~R]mKDNk+DE#@#@&@#@&,P~~kYDU5Vqt+Mn~',J,4+.n,J~[,$EbV[	4+Dnj$VcJ:za+EBJ&J#,[,EPz19PrP'~~Ek^[4+Dj5VvJ&1 p;C	YrYHE~r@*TE*P[~E,bHf,E,[~~EbVNqt.+U;s`rZCOoWMzJBJOgGO_z/kkLxn[r#@#@&,~P,qW~kYD	n(nLP@!@*,JE,K4+x@#@&,~P,P~PkY.j$V4nDP',dODU;^4+.n,[~J,)1GPE~LPA;r^N	t.?5^`r+(nTE~kY.4KL*@#@&,~P,2x9~(0@#@&@#@&~P~~K4%I^WMNjnDRb^Ob\nZKU	+^DkKxP{PkODG4/W	x@#@&,PP,G4NI+1G.NU+DRjW;.1+~',E?AS3/:P$Kr1Y;DDBP,:kDV+Y~,,f/^DbwOrKxTB~$hDW9;^Y;W9+D~~,5ECxDrYHTS~]nDr^TSP]	4KTTBP$&fYSP]?OlDEdD,sI}\PrP[,dOD&Y:dK8s,[~J,EPLPdOM?;s	4+.+,',J~}IG2I,AI~$q+8nTTS~]nDK[E1YZK[nTr@#@&P~P~G(L]+1GD9?nOcSW^V:X2+,x,l[dW13Il96x^X@#@&,P~~K4L"nmKDNUnOR}wx@#@&~~,PdYMKWkYn[&f/~x,JE@#@&@#@&P~,PbxY:WD3xDDr+kPx~Z@#@&@#@&P,Pq6~dYMKHwnn.G1+d/,xPrsE~Dt+U@#@&P~P,~,mC^V,nDKmd/|fCYm`E@!6WDs~l1YkKUxJrRcz^orR(kUzK.NDmWKD: CkwEJ,hY4KN{JJhrUPJr@*E#@#@&~~,PP,^l^VPh.Gm/k{9lOCvJ@!k	2EDPOza+'EE_k[NUrJ~	ls+'rJqn4hoEJ,\Cs!+'r~[,ZC"cfc*PLPsmCd`dYM	+(nLb,[PE ktO:^E,[~;C"`&W#,'Pr@*E#@#@&~~,2x9~q6@#@&@#@&~P,P&0~1GO,W8L"nmKD[jYR36wPPtU@#@&~,P,PP,W(%ImGD9?nOcHW7nsbD/D@#@&P,P,P~P9G,4k^nPgWO~K4L]n1W.NUnDR3}s@#@&P,P,~P,P~PbxOPKY2	ODb+/,x~k	Y:WO2UOMkn/,QP8@#@&~,PP~~,P~PkOMnGkYNqG/,xPkY.nK/On9qfk~[,W4N]nmKD9?nY ob+sNkcJ&fEb,[PE~r@#@&P,~,P~,P,PmmV^~nMW^+k/m(	\+	OWMX{GCOl@#@&,P~P~~,P~PK8L"+^GMN?nOcHG\H6O@#@&iSWKw@#@&P,P~2	N~(6@#@&@#@&P,PP&W~/DD:X2+K.Kmn/k~',JoE,KtnU@#@&~P,~,P~1l^VPhDK^+k/mfmYCcr@!k	2EDPYH2n'rJ_k[NnUrJ~xmh+{JEH(D1nA&Yn:kErP-mV!+'rPL~Z_Ic&W#~',kxDPWD2xD.r+kPLP/C]c2cbPL~J@*Jb@#@&PP~~Ax[P&W@#@&@#@&P,PPK4N]+1W.NU+O ;VWkn@#@&PP,~j+DPK4%In^KD[?OP{PHGDtkUL@#@&@#@&,~,P(6PkYD:XannMW^+k/~x,Jsr~K4+x@#@&~P,P,P~r8%:+aYwUR1VGd@#@&~~,P~P,dY~}4NK+XYwU'	WOtbxL@#@&PP,~P,P/O~r(LwxxxGO4kUo@#@&@#@&P~~,PP~dDDj;^	4+.P{PJ,4nDPEPLP$;bVNq4+M+?$scJ:Xa+E~E(r#~[,EPz19~rP[~$!ksNq4DnU;^`Jq+(KMr~dYMn8ho#,'PrPbg9~J,[,A;ks[qtnDj;^`EjDlY;dr~Enrb@#@&~,P,PP,?OPK4%ImG.9?+D~',?+M-nDcZM+CYn6(LnmDcJzf69~RIn^KD[/Or#@#@&P,PP,P,G4NInmKD[jYRz^Yb\+;GUxmDkGx~x,/ODG8ZKxU@#@&PP~~,P~W(%"+^KD9?+DRUGEMmnP{PEihfb:3PrP[,dOD&Y:dK8s,[~J,j2:P,j:bKijYPxPEpEPE,[,/YM?$s4+.+@#@&~~,PP,~W(LI^GD9?Y SG^0Kzw~',l[JKm3]nmN6x^z@#@&~,P,PP,W(%ImGD9?nOcrwU@#@&PP,~~P,?Y~W8%"+^WM[?Y~x,1WO4bxL@#@&@#@&P~,P,PP&0,dYMnG/D+[(G/P@!@*PrJP:4nx@#@&,P~P~~,P~PkODU;s	4+Dn~{PEPq4Dn,J,[P~EbsNqtnD?5svJqg qGJ~kO.nK/D+[q9d*@#@&P,~P,P~~,P?nO,W8L"n1W.9?YP{PUnD7+.R;DnCD+r(%+1Y`r)9rGAcInmG.9/nYrb@#@&P~~,PP~~,PG4N]mGMNU+Ycb1Ok7+/W	xn^DkW	~',/YM98ZKx	@#@&P~~,P~P,~PK4%]mW.[U+ORUG!D^P{PJ`nG)KAPEPLPdOMqYh/:4V,'~J,?AK~$jPzKi?Y~',BKv,JP'~kY.?$sqtnM+@#@&P,P,~P,P~PK4%]mWM[?YRdG^3:Xa+~'~C9SGm0]+mN6U^X@#@&~,P~P,~,P~K4NI+1WM[?Y ra+U@#@&PP,~P,PP,~j+DPK4%In^KD[?OP{PHGDtkUL@#@&~P,~,P~Ax9Pq6@#@&@#@&,P~PAx[~&0@#@&@#@&A1f,jiAD0QDAA==^#~@%>

<%#@~^mgUAAA==j`A~hDKm+k/|(x7+UYKDzmGlYm@#@&@#@&P,~9qtPbxOpOz@#@&@#@&,~P1lss,nDG^/d{GCDlcrP@!w@*Lx(dwp@!&w@*Jb@#@&@#@&,~P1lV^~KDKm/d{9CDlcJ@!CP	lhn{JP'~;C]`2**P',W(LImK.NU+ORwkns9/`r(fr#PL~/C"`2cbP'~r@*@!zm@*J*@#@&~,PqW~K4%D^KD[k+DR0b+^[/vJKk1Y;.J#,@!@*,JJ,P4+	@#@&P~P~~,mCV^~nMW^nk/{9CDlcJ@!2@*@!rso,/D1P{EPLP/C"`f**P[,G4ND+1G.Nk+DRWkns9/cJhrmDE.nr#P'~;C]`2**P',J@*@!za@*rb@#@&P~PAx[~&0@#@&~P,ml^s~nMW1+d/m9mYC`r@!w@*@!WG	YPdr.+xc@*@!(@*E,[,W4ND^WMNd+DRWrVNkcJ:kY^nE#,[,J@!z8@*@!zWW	O@*,@!8.@*J#@#@&,P~mms^PKMW1+/k{GCYm`G4NDn^KDNknYc0ks[/vJG+dm.raYrW	E#,[~E@!4D@*E*@#@&P,~1ls^PhDW1+kd{GlOlvJK.KNE1OP;WNl~J,[,W8L.n1W.NknYc0rn^N/cEhDGN!^DZG9+r#PLPr@!4M@*E#@#@&@#@&,PP&WPkYD:z2+dkkY~'~E&pEP:4+	@#@&~,PP~~EfGPgGDP9b/aVlHP5;l	YrYH@#@&~,P2^d+@#@&P,~~P,mmVsPK.Kmn/kmfmYCcrpECUDkOX,)7lr^l(V+=Pr~[,W8LM+^GMN/OR6k+^[d`rp!lUYrOHJbPL~J@!4.@*r#@#@&~,P3x9~&0@#@&@#@&PP,mmsV,n.W1+dd|flDC`rnDb^n),jUPEP'~wW.:mOZ!D.n	mXcG(L.+1GMNdYc0kV9d`rn.k1+Eb*P[,EPlm4Eb@#@&@#@&P~P(W,/OD:zwn.G1+/d~{PEsr~Dtn	@#@&PP,P,~mmVsPhDG^//|9lDl`r@!8D@*J*@#@&P~~,P~mmsV,n.G1+/dmGlOlvE@!kUaEDPYHwxJrCrN9+UErPxmh+{JJ&On:rPLPrxOPKY3xD.k/~',ZC]c2cbPL~rP-mV!+'rPL~Z_Ic&W#~',W4N.+1WD9dnYc0b+sNdcrn.W9;mDZG[J#~',ZuIvfW#~LPr@*J*@#@&~P,P~P1lss,nDK^+k/{GCOlvJ@!kUw;O,YzwxJrZ4n134GarJ~xmh'ErpDXJ,[,rxDKGYAxO.b+/,'P;CIvf*#,[,J@*Jb@#@&P~P,~P1lss,nDG^/d{GCDlcr@!6WxDPkry'EJ FEE@*@!4@*)N9PqDnhPDW,r.Nn.@!z8@*@!&0KxO@*r#@#@&~,P3x9~&0@#@&P,PmmV^~nMW^+k/m9mYlvE@!Jw@*rb@#@&@#@&A19Pji~CYYBAA==^#~@%>

<%#@~^tQAAAA==j`A~hDKm+k/|9lDlc/DD9CDl#@#@&@#@&PP,(WPkYMKzwnKMW^+kdP{PEorPY4n	@#@&P,~,P~}4NK+XYwURqDrYSrUP/D.fmYl@#@&~P,2^/n@#@&~,P~P,]+kwGUk+RA.bYnPkOMfCDl,[P74;.S6@#@&P,P3U9Pq6@#@&@#@&2g9~?`AdTAAAA==^#~@%>

<%#@~^4wQAAA==j`A~~EbVN|q	-+	YGDH{).MlX@#@&@#@&PP,~/l^V,A;ks[GlOlz.DmXcE1W:2C	Xm/O!wE*@#@&@#@&,P,~/DD9lDl~x,?+M-+MRHm22lDtvJwkU^^E[+kwJ*P'~r-kU-xOWMz|kO:kRk	mr@#@&,P~PU+O~K4LwUP{PZMnCYr(LnmOcr?^Db2YbxL wkVnjH/O+s6(Ln1Yr#@#@&P,~PU+OP}4%P6YwUP{Pr(%oxcZM+CYnP6Osbs+v/O.GlYCS,Y.Eb@#@&@#@&P,PP}4NP+XYoxc.rD+SbU+,J@!r~'P;C"`fGb@#@&P~P,64NKnaDsx 	MkO+dr	+~rf&HPmDH(Y:d`*J@#@&,PP,64NK+XOoxcMkO+Jr	+~J&WP"+5;/Y }!+.XUOMkUT`rJZrJ*~@!@*PEJrJ~P4+x,E@#@&PP,~64NK6OsU qDrYJk	+~E,PPj3U?(rgcrJ/rJ*P',I5E/OR5En.H?YMrxT`Jr/EJ*J@#@&~P~~}4%KaYwx 	MkYnJbxnPr3	N~&0r@#@&@#@&,~P,r8L:+aOwxRq.kD+SbUnPrs`1/K(6gP!+D(x7+UOKDXE@#@&P~P,6(LP6DsxcMrYSrxPEjV+1OP;b?A~j2U?&rH`EE;JE#r@#@&@#@&~~,PsG.,kUYU;(Px,F,YW,k	OKKY3xDDrnk@#@&,~P,PP,(WPbxD?;4~x,F~K4nx@#@&~~,PP~~,P~r(%:+aDs	RMkDnSbxnPrZ)jAPJrEJBPJ,'~Z_Iv&*#~',kUYU;4,[~/_I`f**@#@&P,~,P~,2^/+@#@&,~P,P~P,P~6(LKaYwxRq.rYSbxnPE/z?3Pr~[,Zu]v&cb~LPrxDj!4~LP;CIv&Wb@#@&P~P,P~~AxN,(0@#@&P,~~P,PU2j?(6g`EZrbP{PrUD?E8@#@&P~P,~,P/mV^P?Y!2{;Whwmxzm&x0K@#@&,PP,~~P1l^V~fCOm4C/m?Y;2v/Y./K:2l	zglh#@#@&P,P,~P,mCV^P$;bVN|)DMlX@#@&~P,Pg+aY@#@&,P~P}8L:+aOwxR	.bYnSbUPEA1GP?V^Yr@#@&P,P~6(LKaYwxRq.rYSbxnPE3gf~s`HZ:q6Hr@#@&~~,P64NP6OwxcDbYJk	+~Z_IcfF#PL~J@*J@#@&~~P,/Y~r8%:+aYwU'	WO4bxo@#@&,P~PknDP6(Lwx'	WD4k	o@#@&@#@&3HGP?`$YlQBAA==^#~@%>

<%#@~^qAUAAA==j`A~~EbVN|bM.lH@#@&@#@&P~~,fqt~k	Y?!8@#@&@#@&,P~PjnDPG4N]+1W.[U+Y~x,?nD7nMR/M+mY+}4NnmD`EbGr9$cI+1GD9/+DEb@#@&P,P~/O.U;s4nDPx~rP4nM+~J,',A;bV9tDj;^`EKHwnEBJqrb@#@&PP,~dYM?$V	tn.PxPkODU;s	4+Dn~LPEPzHGPE,[,AEbV9	tDn?$VcE&1R5;l	YkDzE~r@*ZJb@#@&~,P~W(%ImG.9?+O zmOk7n;WU	+1YkKx,xPkY.f(ZGU	@#@&,~P,W4N]nmKD9?nY jKE.m~',Jj3d2ZP~]KrY^nY~~]nMWN!mD/W9+D~,$K.bm+Y~s"rH,E~[,/DD(YnhkK8V,'PrPE~LP/O.U;s4nM+~LPrPr"fA]P~5~$hDG[!mY;GNTJ@#@&~P,PK4%In^KD[?ORdW^V:Xwn~{PCNdG13]l9rx^X@#@&P,P~W(L]n1WD9j+DRranU@#@&@#@&P~P~(6PHWD~W(L]n1WD[jY 2}o,K4x@#@&P,P,~P,kUYUE8~{P!@#@&P,PP,~~W(L"+^W.[U+ORtG\sr.kY@#@&~,P~P,~GW~qtbV+,1KOPK4%ImG.9?+D 2}s@#@&~~P,P,P~P~r	YjE(~',kUOUE4~Q,F@#@&,~,P~,P,PPK4N]+1W.NU+O tW\H+XY@#@&7JWKw@#@&~P~~,P~W(%ImG.9?+O tW-+wrM/O@#@&,PP,P,~r(LP+XYoUcDbO+dkx~EIfb:~l.z&Yn:kcJ,[~r	Y?;8,[~JBf*J@#@&P,PP,P,64NKn6DsU qDkDnSbx+,ECDHqD+h/cTB!bP{~J,[~r	Y?;8@#@&~P,~,P~bxD?E(P{~!@#@&~P,P~~,fW,	tbV+,HGY,W(L]+^GMNj+D 2}s@#@&,PP~~,P~P,r	Yj!4,'PbxDjE(PQP8@#@&~,PP,~P,PP}8%K6DsUR	.bYnSbU+,JC.HqYnhk`EPL~bxOUE(P[,JBq#,'~J,[~/_I`2*#,[PK8%ImKD[?nOcsr+^[/vJPrDV+Eb,[~Z_]v&**@#@&PP,P,~P,P~r(LPnXYs	 MkYJrxPrl.X(O:d`r~[,kUOUE4~',JS *~{PE,[,ZC"`2*#,[~W(L]n1WD9j+DRsbnsNk`rn.W[;1Y/W9nJ*P'~;CIcfW#@#@&,~,P~,P,PP}4NP+XYoxc.rD+SbU+,JlMz(Y:k`EP'~bxO?!8PLPES2#Px~rP'P;u"`fW#,[PK4N]+1W.NU+O wk+^[/vJnMr^+r#,[~Zu]v&*#@#@&P,P~~,PP~~K4%I^KD[U+DRHK\H+XY@#@&iSGGa@#@&,~P,2x9~(0@#@&,P~PG8NInmK.NU+O ;VWdn@#@&~P,~U+O,W(LImK.NU+OP{PHGDtk	L@#@&21G~jj~PqY0BAA==^#~@%>


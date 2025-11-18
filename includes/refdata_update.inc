<%
'          **************************************************
'          *  Copyright Fulcrum Computing, Inc.  2000       *
'          **************************************************
'          * This program is copyrighted under the laws of  *
'          * the United States. Any unautorized duplication *
'          * or use of this progrm is prohibited.           *
'          **************************************************
'**Start Encode**#@~^BAAAAA==@#@&FwAAAA==^#~@%>

<!--#INCLUDE FILE="../includes/write_errors.inc" -->
<!--#INCLUDE FILE="../includes/db_setup.inc" -->
<!--#INCLUDE FILE="../includes/refdata_tbl_setup.inc" -->
<!--#INCLUDE FILE="../includes/transfer_html.inc" -->
<!--#INCLUDE FILE="../includes/get_refdata_db_info.inc" -->
<!--#INCLUDE FILE="../includes/logon_check.inc" -->
<!--#INCLUDE FILE="../includes/get_program_path.inc" -->
<!--#INCLUDE FILE="../includes/pad_data_string.inc" -->
<!--#INCLUDE FILE="../includes/build_where_sql.inc" -->
<!--#INCLUDE FILE="../includes/company_setup.inc" -->
<!--#INCLUDE FILE="../includes/display_edit_errors.inc" -->

<%#@~^0hgAAA==@#@&@#@&,P,fqtPkODdWG3`wFnH~P(GV"+mwU[~,/DD94)^DkGx@#@&P,P9(tP/O."+5|zBPr	YADDg4M@#@&,P~f&H~dDDZ4V~,kxDH8DAxDDr+dS,4GV"nYMkn-@#@&~~,f(H,CMXjV"+;k`*SPo~~qBP((BP/D.K:wB~dYMb1mn/dJ7V@#@&,~PGq\~kYDunmNrxTP6OB/DD?$Vq4+M+SPkY.j$V4nD ~,dODhosSGm@#@&@#@&~P,dYMnLhdWm~x,MnYhLsnCDtvJJ*@#@&@#@&,P~ZmVs~U+Y!2{;W:aCUX|q	0G@#@&~,P^l^sPGlOC(l/nmU+OEackY.;Wswl	XgC:#@#@&@#@&~~,q0,]+$E+kO sKDs`E^3aDJbP@!@*Pr5E~:t+U@#@&P~P,~&0~"+$E+kYcoWM: ZKEUO,'PZ~K4+x@#@&~P,P,P~PdOMKn:a~',Ji]cJ@#@&~,P~PAsk+@#@&P,PP,P,~/DDP+sw~x,JJ@#@&P,PP,3UN,q6@#@&P~~,P^l^sPdWLG	{Z4n13c/D.:+haPLPJM+6[lDlmEaNCORlk2JB/YM/G:al	XHlhn*@#@&P,~P,/O.zmmndkS-V,x,?3U?&r1vJz^m/dS7VEb@#@&P,~2^/+@#@&~P,P,/OD)^1+d/d-V,'~E"J@#@&~,P3x9~&s@#@&@#@&PP,?OPD.Hkod~{PZMnlD+r(%nmD`r?^Dr2DkUoc9k1YrG	lDzE*@#@&@#@&~,P(6P"+;!+kORwW.:vJ\n	EJ*~@!@*PJr~Ptx@#@&~P~~,P^l^sP"+O;Mx{PG|Hnx!@#@&P~,2	NP&0@#@&@#@&P~P&0~];EdYcsWMh ZKE	Y~@*~T,lUN,]+$EndDRsG.s`EHz(gK]ApUJ#,@!@*~JIJ~Y4+U@#@&PP,~P,kxDH8DAxDDr+d~{PT@#@&~P,P~~wWD~3mm4PX~&x~"+$E+kYcoWM:@#@&,P~~,PP,(0,I+$;n/DRwW.:ca*P@!@*,EJ,lU[,I+5;/ORwGM:cX#,@!@*,JKUJ,lUN,In5!+/D sKD:vabP@!@*,JjE8hbYEPmUN,In5!+/O wW.:va*P@!@*Prjw9lDnJ,lUN,6~@!@*PJ;G:alxHE~K4+	@#@&P~~,P~P,~P,kUOg4D3UDDr+k~{Pr	Yg4DAxD.k/~_,F@#@&,PP,~P,P2	[~q6@#@&P~P~~,P~q6~6,'~E"+Y.r\nJ,P4+U@#@&,PP,P,~P,P~/DD]n$|+H~',I+$;n/DRwW.:cEfIn0&9J*@#@&~,PP~~,P~P,8KV]YMk+7+,xPDD;+@#@&~~,PP,~PAV/(WPXP{PE^]n6q9J,Ptx@#@&,PP~~,P~P,~U2jUq}1`r^"n;"+WqGJb~{PI5E/YcoGDs`X#@#@&~~,P~P,~2	N~(6@#@&~~,P~PgnXY@#@&P,PP,PbrP{P.+$EndDR0K.:cmW!UO@#@&P,P~P~dDD94z^YbWU~{Pj/Ck+cD5!+dDR6WDs`br#*@#@&P,P~~,k0,dYMf4z^OkKx,'~Jji~H(Kr~K4+U@#@&PP~~,P~P,dDD9(b1YkKx,xPrb9fr@#@&~,PP,~2	NP&W@#@&,P,2s/n(6P]+$;+kY }!+DzdDDrxTcrIn6qGJ#,@!@*~JrPPtx@#@&,PP,~PkYD"n5|X,'~In5!+dYc}EDzdDDkULvJ]+6(GJb@#@&,PPAx9~q6@#@&@#@&P~~&0P"n;!+/D oWM:vJ\b(H:I3pUE#,'~EIJPP4x@#@&,~,P~,IfksPm.XU+sI;dc"+;!n/DRsK.hR;W!xO#@#@&,P~P,~kbPx~Z@#@&~~,P~PwGMP3mm4P6,q	~I;;+kY oKD:@#@&P,PP,~~PEnMW^+dd,lsV,92d2P3,I+5;/O/,obDdD@#@&PP,P,~P,qWP^+WOv6~8bP{PJGE~K4+	@#@&P~~,P~P,~P,4Gs"+moU9PxPVnDIn6f~qx6Wvhk9`a~y#b@#@&PP,~P,PP,~~q6P(WsIn^wx[P{~YMEn~:t+U@#@&P~P,~,P~,P,PP,q6~/DDj;^4nM+P{~JrPK4nU@#@&P,P~P~~,P~P,~P,P~dDD?5sqtnD~{PE,4+DPr@#@&,P~P,P~~,PP,~P,PPkO.?$VqtnDn+,'~J,	tDn~rP[~$!ksNq4DnU;^`J:XanJBVn0D`G8NI+1GD9?+D okV9/cJHCs+E#Bq#*P'~rPbH9,J@#@&,~,P~,P,PP,P,~2^/n@#@&P~~,PP,~P,PP,~~P,/DDj;s	4+.+,xPkY.j$V4nM+~[,E,r],J@#@&P,P,~P,P~P,P~~,PPkODU;Vq4nD ,'~/O.U;s4nD ~',JP6],J@#@&,~,P~,P,PP,P,~2	N~q6@#@&~,PP,~P,PP,~~PkYM?5V	4DnP{~/DDj5^tn.P'P~;bV[qtD+U;^cJ&1 qGJShbN`XS *#@#@&~~P,P,P~P~~,P~/D.?$V	4D++~{PdYMj$V	4+M+ ,[,$EbV[4+.nU;VvG4NI+1G.NU+DRokns9/cJ:zwJbSK4L]n1W.NUnDRob+^N/vJjCV!+E#*@#@&~,PP,~P,PP,3UN,q6@#@&P~~,P~P,3x9P(W@#@&P~~,P~1aD@#@&,P,PP,q6~/DDj;^4nM+P@!@*PrJP:4nx@#@&,P~P~~,P~?OPK4%]mW.[U+OP{~U+.7+MRZM+mO+}4%+1YcEzfrG$R"+mK.[/Yr#@#@&~~,P~P,~PK4%]mW.[U+ORz^Dk-ZKxxmDrW	PxPkY.9(ZW	U@#@&PP,~~P,P,W8L]n1W.NUnYc?G;Mm+~x,J92d3:2~wI}HPrPL~/DD]+6K8s,[Pr~J,[PkO.?$VqtnDn@#@&P~P,~P,P~G(LIn^KD[?OcSG13:XwP{~l9SGm0InC9rx^z@#@&PP,~~P,P,W8L]n1W.NUnYcr2n	@#@&~~,P~P,~,?nDPK4L"+1GD9?nY,'~HKYtbUo@#@&P,~~P,P,P^lss,j2NmO+|qOns/{PC(Vn`rEBJEBJ@$J[kYMj;^4+M++Sr1WD~bk/kTUnNr#@#@&~P~~,P3x9~q6@#@&~,PP~~wW.PAC1t~XP&xP"+$;+kY sKDh@#@&PP,~P,PP&W~V0D`a~qb,'~J`EPmx[~"+;;nkY sK.s`EGJ,[Psk9c6B b#,'~ErPK4nx@#@&P,~~P,P,P~Prr,'~kb~_,F@#@&,PP~~,P~P,~mDzU+^I+$/vrk*PxPsk[cX~ *@#@&,PP,~~P,2	N~qW@#@&P~P,~Pg+aO@#@&P~~,P~q6~bk~@*PZPK4+	@#@&,P~P,P~~,?2Ujq}1`rC.X"+$/E#~x,l.XUnV"+5d@#@&P~~,P~P,~U2jUq}1`rKKOl^In;kJb~{Pkb@#@&,PP,~~2^/@#@&P~~,P~P,~BgW~iaNlOn,In;!nkYd,HmN+@#@&,~P,P~P,P^C^VP"nY!Dx|PG{t+	E@#@&~~,P~PAUN,qW@#@&PP~3	N~q6@#@&@#@&,P,q0,?Aj?&rH`rKGOmVI5/r#P@!@*~JrPmx[PdOMf8b1OkKx~@!@*PJiKGbP2r~:tn	@#@&PP,P,~Ifr:,l.zU+V"n;k`?Ajjq}1vJPWOC^In;kE#*@#@&~,PP~~kY.I5n+z{?A??&rgcJmDzI;dE*`F*@#@&,PP,~~q6PU2j?(6g`EKKOl^In5kJ#~@*,F~K4n	@#@&,P,PP,P,~sKD~kbPx~yPYK~?A??&6H`rKKYCV]n$/E#@#@&P,P~~,PP~~,PCDHjV];k`kbO8bP{Pj2U?(6g`Jm.X"+;kEb`bk*@#@&P~~,P~P,~16O@#@&PP~~,P3x9~&0@#@&P,PP,PU3?Uq61vJPGDlV"n;kJ#,x~?A?Uq61cE:WOl^]+$/Eb,OPq@#@&P~P,~,qW,?A??&rgcJ:WOl^In5kJ#,@!P8PK4nU@#@&P,P~P~~,Pj2Ujq}1cE:WYCs"+5/rb,'~rJ@#@&P,P,~P,P~?A?j(}1`rCDHI+$dE#,',JE@#@&~,P~P,3Vk+@#@&,PP~~,P~PU3U?(}1vJlMX"n;kJbP{PC.H?+^]+$/@#@&~~P,P,2UN~(6@#@&P,~2^/n(6P?3jUq61vE"+5n+HJ#,@!@*~JrPPtx@#@&,PP,~PkYD"n5|X,'~?3jUq61vEI;FnHJ#@#@&,P~P,~U2jUq}1`rI5|XE#,'~Er@#@&,~PAxN,(W@#@&@#@&P~P(W,/OD"n;n+z~@!@*PEE,Y4+	@#@&P~,P,/YMSKG3`wF+HPx~kYD"n;n+X@#@&~P,P,4GV]n1sUN,xPV+O]0f$(	0G`kOMSGK3`w|X*@#@&,P~P,qW~(WV"nmwxN,x~YMEPPtnU@#@&~P,~P,P~(6P4Gs"+^s	[,'~DD!+P:tU@#@&P~P,P~~,PP,dYMf4z^OkKx,'~JiKGbP2r@#@&,P~~,PP~~,Pj2Uj&rHvJfI+$IWqGJbP{PdOMSWKVja|+H@#@&P,P,P~P~3^/n@#@&~P,P~~,PP~~"+9ks~mDz;EMI+6I^fmYC`W!b@#@&PP,~P,PP,~~/DDG4)mOrKx~',EbGfE@#@&PP~~,P~PAU9P(6@#@&PP,P,~P,ZCV^P9rkwVmz{wWDs@#@&P,P,P3Vdn@#@&~P,~P,P~/mVV~3MDGD|\/dmo@#@&,P,~PAx[P&0@#@&,PPAs/q0,]n;!+kY sG.sR/W!UY,'~T,KtnU@#@&~P,~,InGksPlMX;;D"+WIm9CDl`WT#@#@&P,~~PkYMf8b^ObWUP{~Jzf9E@#@&P~~,P/l^s,frkw^lX|sK.:@#@&~P,2sd@#@&,~P,Pq6~dYMf(b^YrG	PxPrinGbP3rPK4n	@#@&P,~,P~,/DDSKW0iwn+zP{Pj3U?q}H`r^I5]+6qGJb@#@&~,P~PAs/@#@&~,PP~~,/ODdGK3ia|XP{P"n;!+dYcsG.s`Jf]+6qfrb@#@&,P,P~2U[,qW@#@&@#@&,P~~,q0~dDD94z^DkG	P@!@*Pr?n(nrPPtx@#@&,PP,~P,Pq6~nDMHkodR/G!xOP{~!,lU[,`/O.zm^+kdd\s,',JjrPK.P"+5E/O wWDscJf26DEbP{Pr5E#~P4+U@#@&~P,P~~,PP~/mVsPjC^k[mY{q	w!O@#@&P~P,P~~,PP&WPDDtdL/cZKEUY~@*,!~K4nx@#@&~~,PP~~,P~P,~mDzADMWDt/TdP{PnDMHdLkRqDn:k@#@&,~~P,P,P~P3U9P(0@#@&P,P~~,PP~~D.R1sl.@#@&,PP,P,~PAx[P&0@#@&@#@&P,~P,PP,(WPDMHdod ;W;xD~',!~C	NP8G^2.DwU9P@!@*PDDEP:4+	@#@&P,P~~,PP,~P&0PkO.f(b1YrWU~{PEbG9J,K4n	@#@&~~,P~P,~,P~,P,ml^V,)N9{]+1W.[@#@&P,~P,PP,~~PAVk+@#@&~~,P~P,~P,P~~,mlss,j2NmO{]mKDN@#@&,~P,P~P,P~~AxN,(0@#@&P,~~P,P,2s/n@#@&P~P,~P,P~~,Zlss,2.DK.|Hnk/mo+@#@&,~P,P~P,2U[,q0@#@&P,PP,3UN,q6@#@&@#@&~,P~PUnY,W8%"+mG.9?nY,x,1GDtbxo@#@&@#@&P,P~P&0~8KV2M.s	NP{~OD!+,K4+U@#@&P~P,~P,P]nkwWUdR3x9@#@&P~,P,2Vk+&WPD.Hkod ;WE	OP@*P!,P4+	@#@&P~P~~,P~mmsV,frdaVlzmAD.WMd@#@&~,P,P2^/@#@&,P~P,P~oKDPACm4P6,(UPU2U?(rH ;rHKAHKU@#@&~,PP~~,P~P&W,Vn6Yv6~8#,xPr^EP:tnU@#@&P,~P,PP,~~P,PU2j?(6g`a#,xPrJ@#@&,PP~~,P~P,3	N~&0@#@&P,P,~P,1n6D@#@&~,PP,~PkYD:nhw,',JE@#@&@#@&P~P,~P,qW~kYD)^1+d/d-^Px,J`JP:tU@#@&P~P,P~~,PP&WPU2?U(61vJ:WOls];dJ*~@!@*PEE,WD~dDD94z^DkG	P{PJzfGEP:tnx@#@&~~,PP,~P,PP,~^l^V,.kOn|K.l	d0Dmu:HScdDDKosJKm~LPrD+6NmOl|E2NmYn m/wrb@#@&PP,~~P,P,P3Vdn@#@&~P,~P,P~~,PP~^mVsP"nDE.	{:W{t+	;@#@&P~P,P~~,PPAUN,q0@#@&~P,P,P~2sd@#@&P,~P,P~~,PmCs^P]+D;Mxm:W|H+	E@#@&P,P~P,P3U9Pq6@#@&,PP,~3x9P&0@#@&~~,2UN,(0,QzUGAA==^#~@%>

<%#@~^XgIAAA==@#@&@#@&Uj~P.mVb[lD+mq	w;O@#@&@#@&~PEzz,#CVbNmYnPWGM:~+	ODb+d @#@&@#@&~,PrxD3MDH(D,'PZ@#@&@#@&,P~q6PdOMf4z^YbWx,x~JzfGJ~K4n	@#@&P,~P,P(W,YDrhvIn;!nkY wWM:`rKH2+r#bP{PEE,KtU@#@&PP,~~P,P,kUY3.M18D,xPbxO3MD18.,_~F@#@&,P~,P,PP,+M.HkodRzN[~bxYA.Dg4DB~EKHwP]+5;bDnNr@#@&,P~~,P2U[,qW@#@&~,P~,P&0PDDbh`"+5E/O wWDscJgl:Eb#,',JEPP4x@#@&,~P,P~~,PkUOAD.1(.,'~bxD2DM1(.P3Pq@#@&P~~,PP,~PDDtdL/cb9N~kUOAD.1(.~,JHCs+P]n$ErD[r@#@&,P,PP,2	[P&0@#@&,P~3	NP&W@#@&PP,(WP"+$En/O wW.:vE.mV;nr#Px~rJ~K4n	@#@&,P,PP,k	O2MDH4MPx~bxYA.Dg4D,Q~F@#@&,P~P~~D.HkL/cb[[,kxO3MDH4MS,J#mV!+P"+$;kM+[J@#@&~~,2x9~q6@#@&@#@&3x9PUj$@#@&@#@&N5cAAA==^#~@%>
<%#@~^AwMAAA==~Uj$,fb/w^lHmsKDh@#@&@#@&~,Pf&\PkYDzsro	~,/ODurTtskT4Y;V.@#@&@#@&~~,qWPkOMf8zmDkW	P@!@*Prb9frPP4x@#@&~P,PP,(WP(W^InY.r\nP{~YMEn~KDPdOMb^mdkS-^P@!@*Prjr~K4+U@#@&P~~,PP,~PkYD_nCNbxTKn6O~{PE@!6GxDPdr.+'EE3FEJ,^KVGM'rJawswoswJE@*;E..xY,(x6WDsCOkKx@!zWWUO@*J@#@&,~P,P~3^/+@#@&,P~P,~,P~kYMC+mNbUo:+aY,'~E@!0W	OPkkyxEJ3FrJ~mGsKDxJr:swsoowJJ@*~tlV+,];;bDNP;tmUo/E@#@&P~~,PPAUN,q0@#@&~P,2^/n@#@&~,P~P,dYMCnC9kxLP6OP{~r@!WKxDP/byxJr_qJrP^G^WD{EJ[sswoosrJ@*P)N[~g+APAUYMXE@#@&PP~3	N~q6@#@&@#@&,P,q0,VU`kY.Cl[r	oKaY*P@!{~0*,WMPdY.9(b^YbGx,@!@*~rbf9E,K4+	@#@&P~,P,P/DDzskTx~',JCsbox{vZA1KA]vJ@#@&,P~2U[,qW@#@&@#@&,P~(6P/O.zm^+kdd\s,@!@*PJ`J,Cx9PdYMf8)1YkKUP@!@*Pr)9frP:tnx@#@&,P~P,~/DDurTtVrL4Y/VM~{PEr@#@&PP,2^d+@#@&~P,P~~kYD_ro4VkT4OZ^D,'~J:$G!T!ZE@#@&P~~AxN~(6@#@&@#@&ZcQAAA==^#~@%>

   <html>

   <head>
   <title><%=#@~^DgAAAA==dDD/K:alxH1mh+sQUAAA==^#~@%> Reference Data Input Form</title>
   </head>

   <body bgcolor="#008080" >

   <div align="CENTER">
   <table border="1" cellspacing="0" cellpadding="2" bordercolor="#000080">

      <tr>
        <td <%=#@~^CQAAAA==dDD)^kTxPZAMAAA==^#~@%> valign="TOP" bgcolor="#000095">
        <b><%=#@~^DwAAAA==dDDul9kxTKaY,zgUAAA==^#~@%></b>
        </font></td>
      </tr>

      <tr>
        <td valign="TOP" bgcolor="#B4B4B4">
        <font face="Arial" size="2">

     <div align="CENTER">
   	 <table width="90" border="0" cellspacing="0" cellpadding="5">

        <form Action="http://<%=#@~^CQAAAA==dDDKT:dWmmwMAAA==^#~@%>refdata_update.asp" Method=POST>
          <%#@~^JQAAAA==(6PjA?Uqrg`ryI;]+6q9E*P@!@*~JrPK4nUPSwkAAA==^#~@%>
        <input type="Hidden" name="$RefID" value="<%=#@~^FQAAAA==jA?j&rg`JfI5I0(fr#~zwUAAA==^#~@%>">
        <input type="Hidden" name="$OrigValue" value="<%=#@~^IAAAAA==CMX/!D"+0"+19lDlck	Y]n6.l^;+"+m*9wsAAA==^#~@%>">
          <%#@~^BwAAAA==3	N~&0,BgIAAA==^#~@%>
   	    <font face="Arial" size="3">
        <tr>
          <td nowrap>
          <%#@~^GwAAAA==(6PdDDG4b1YbGx,'~Jzf9E,KtUZQgAAA==^#~@%>
          <font face="" color="<%=#@~^DwAAAA==dDDubo4VkTtD/VMEgYAAA==^#~@%>">
          <%#@~^BgAAAA==3	N~&05gEAAA==^#~@%>
          <b>Type</b>
          <%#@~^GwAAAA==(6PdDDG4b1YbGx,'~Jzf9E,KtUZQgAAA==^#~@%>
          </FONT>
          <%#@~^BgAAAA==3	N~&05gEAAA==^#~@%>
          </td>
          <td nowrap>
          <%#@~^5wAAAA==(6PdDDG4b1YbGx,'~Jzf9E,KtU@#@&PP,~~P,P,P~P(W,l.X;;D"+W]mfCOm`rxD]0PHwI+1#,xPrJ~K4+U@#@&PP,~P,PP,~~P,P,/ODPnsw~',j2U?(6g`J]Pr#@#@&,~,P~,P,PP,PAs/@#@&P,P~~,PP,~P,PP,dOD:+sw~'~CMX/EM]+6In^GlYCcbxOIW:X2Im#@#@&,~P,P~P,P~~,2x9~q6NTkAAA==^#~@%>
          <input type="Text" name="Type" size="30" maxlength="30" value="<%=#@~^BwAAAA==dDDP:a7wIAAA==^#~@%>">
          <%#@~^BAAAAA==3^/niQEAAA==^#~@%>
          <%=#@~^IAAAAA==CMX/!D"+0"+19lDlck	Y]n6KXanIm#,vAsAAA==^#~@%>
          <input type="Hidden" name="Type" value="<%=#@~^HwAAAA==CMX/!D"+0"+19lDlck	Y]n6KXanIm#nAsAAA==^#~@%>">
          <%#@~^BgAAAA==3	N~&05gEAAA==^#~@%>
          </td>
        </tr>
        <tr>
          <td nowrap>
          <%#@~^GwAAAA==(6PdDDG4b1YbGx,'~Jzf9E,KtUZQgAAA==^#~@%>
          <font face="" color="<%=#@~^DwAAAA==dDDubo4VkTtD/VMEgYAAA==^#~@%>">
          <%#@~^BgAAAA==3	N~&05gEAAA==^#~@%>
          <b>Name</b>
          <%#@~^GwAAAA==(6PdDDG4b1YbGx,'~Jzf9E,KtUZQgAAA==^#~@%>
          </FONT>
          <%#@~^BgAAAA==3	N~&05gEAAA==^#~@%>
          </td>
          <td nowrap>
          <%#@~^5wAAAA==(6PdDDG4b1YbGx,'~Jzf9E,KtU@#@&PP,~~P,P,P~P(W,l.X;;D"+W]mfCOm`rxD]0Hm:I+1#,xPrJ~K4+U@#@&PP,~P,PP,~~P,P,/ODPnsw~',j2U?(6g`J]Hr#@#@&,~,P~,P,PP,PAs/@#@&P,P~~,PP,~P,PP,dOD:+sw~'~CMX/EM]+6In^GlYCcbxOIWglhIm#@#@&,~P,P~P,P~~,2x9~q67TgAAA==^#~@%>
          <input type="Text" name="Name" size="30" maxlength="30" value="<%=#@~^CAAAAA==dDDP:aPDwMAAA==^#~@%>">
          <%#@~^BAAAAA==3^/niQEAAA==^#~@%>
          <%=#@~^IAAAAA==CMX/!D"+0"+19lDlck	Y]n61lsnIm#,mwsAAA==^#~@%>
          <input type="Hidden" name="Name" value="<%=#@~^HwAAAA==CMX/!D"+0"+19lDlck	Y]n61lsnIm#ewsAAA==^#~@%>">
          <%#@~^BgAAAA==3	N~&05gEAAA==^#~@%>
          </td>
        </tr>
        <tr>
          <td nowrap>
          <font face="" color="<%=#@~^DwAAAA==dDDubo4VkTtD/VMEgYAAA==^#~@%>">
          <b>Value</b>
          </td></font>
          <td nowrap>
          <input type="Text" name="Value" size="30" maxlength="255" value="<%=#@~^IQAAAA==CMX/!D"+0"+19lDlck	Y]n6.l^;+"+m*~FwwAAA==^#~@%>">
          </td>
        </tr>
        <tr>
          <td colspan="2" nowrap>
          <hr>
          </td>
        </tr>
       <tr>
       <td colspan="2" align="center" nowrap>
       <%#@~^DAEAAA==(6PdDDG4b1YbGx,'~Jzf9E,KtU@#@&PP,~~P,P,P~/O.:+hw,xPr?;8skYE@#@&P~P,~,P~,2^/+&0,dYMb^m/dJ7VP@!@*PrjJ,P4+	@#@&P~P~~,P~P,~/DDPnswPx~rZGxDr	Enr@#@&PP,P,~P,P3Vk+@#@&,PP,~P,PP,dOD:+swxP/u"`fc*~[,VnWD`/O.G4)mDrKxS8#,[PdZmd+v:rNv/O.G4b1OkKx~ybbPLP;C]`f**@#@&P,~P,P~~,2x[~&0~50EAAA==^#~@%>
          <input type="Submit" name=<%=#@~^IQAAAA==~;C]v&W#PLPkODG4)mDkGU,[P;uIv&c*~kwgAAA==^#~@%> value=<%=#@~^CAAAAA==dDDP:aPDwMAAA==^#~@%>>
          <%#@~^VgAAAA==(6PdDDG4b1YbGx,@!@*Prb99rPl	[PkYDz^^+k/d\sPx~rjEPmUN,?3jUqrHcrKGYms"+5kJ*P@!@*PrEP:tnxzRgAAA==^#~@%>
          <input type="Submit" name="Skip" value="Skip">
          <%#@~^BgAAAA==3	N~&05gEAAA==^#~@%>
          <input type="Submit" name="Menu" value="Menu">
       <%#@~^MgAAAA==(6PdDDG4b1YbGx,'~Jzf9E,lx9~+MDHkLdR;W!xOPx~ZPOtUP3w8AAA==^#~@%>
          <input type="Reset" value="Reset">
       <%#@~^BwAAAA==3	N~&0,BgIAAA==^#~@%>

       <%#@~^NwAAAA==(6PnMDt/okR;GE	Y~',!~C	NPkOD_ko4sro4Y;V.P@!@*,JEP:4+	P~CBEAAA==^#~@%>
          <br>
          <b>Required Fields in<font face="" color="<%=#@~^DwAAAA==dDDubo4VkTtD/VMEgYAAA==^#~@%>"> Red</b>
          </font></td>
       <%#@~^BgAAAA==3	N~&05gEAAA==^#~@%>
        </tr>
        </font>
        </form>
        </table>
        </div>
        </td>
      </tr>
   </table>
   </div>
   </body>
   </html>

<%#@~^CQAAAA==~Ax[,?`APYQIAAA==^#~@%>


<%#@~^EwEAAA==@#@&?i~PADDKD|\+k/Co@#@&@#@&PP,(0,2DM HEs4D~'~RyF*GyqGRGf~:t+U@#@&P~P,~,kUD2MD1(D,xPbxO2MDH8MP_,q@#@&PP,~~PDMHdod zN[PbUYAD.H(D~~E&f~b^.l[HPA6kkYkE@#@&P~PAx[~&0@#@&~P,q0,n.Dt/T/ ZG;	Y~@*,TP:tnU@#@&P~~,P~lMzAD.KDt/okP{~+MD\/T/ (D+:k@#@&,PPAU[P&0@#@&~P~3MD Z^nlM@#@&3	NPj;(LEcAAA==^#~@%>

<%#@~^9QEAAA==j`A~`w9lY{"nmKD[@#@&@#@&~,P4KsIms	[~',MY]+W9~qU0Kc/DDJGK3j2FXb@#@&@#@&P~,q6P4KV"nmwx[P{PO.!+P:4+	@#@&,~~P,W	PnD.GMP.+k;:PUnXY@#@&~,P~P;C^V~`w9lY{GCYm4C/@#@&~,PP,(0,2DM HEs4D~@!@*~ZPPtU@#@&P~~,PP~~1lsV,	MkO{ADDKDk@#@&,P~P,2sdq0,]+$E+kO sKDs`E^6.bo#l^;+r#~@!@*PIn5!+dYcoKDhvJjlV!+rbP:tnx@#@&~~,PP,~P1lV^~iw9lD+mqOns/mKm8V`sn6Y`]n$En/D wW.s`r1ls+rb~8#SI;;nkYRwGDs`J:z2+r#BIn;;nkY sK.:vJy6Mko#C^EnJ*S"+5!+kYRwWMh`r.CV!+Eb*@#@&,~P,P2	[~q6@#@&P~P~~K4%I^WMNjnDRZsGk+@#@&,~,2U9P&0@#@&@#@&31GPjj~OYwAAA==^#~@%>

<%#@~^0QEAAA==j`A~zN9{ImK.N@#@&@#@&,P~jYPK8L"+mK.[?Y,'~?n.7+.R;.+mYn6(L+^OvJ)f}9~R]mKDNk+DE#@#@&~P,W8%"+mK.NU+Yc62+	PkY.InW:4s~,dYMf8/KxxS~mN6wUUYCDk1~PmNdGm0nn/kkhrkYk1SPmNZs[Pl(V@#@&P~~Kx~+M.WMP.nkE:n~	+aY@#@&,P~K4NI+1WM[?Y b9NHnS@#@&,~P;lV^~iw9lD+mfCOm4C/@#@&,P~(6P2..c1;:(nMP@!@*PZPK4+	@#@&,P~P,qW~ADDcHEs4+M~xP  8c{ q{RGfP:4+	@#@&~,PP~~,P/l^s,2.MWM{H/kCo@#@&P,P~~AV/@#@&,PP,~~P,mmVsP	.bYn{A.DKDd@#@&PP~~,2UN,(6@#@&,P,2x9P&W@#@&P~PK4%]mWM[?YR;sG/@#@&@#@&2H9,?iAFoEAAA==^#~@%>

<%#@~^iAIAAA==j`A~`w9lY{GCYm4C/@#@&@#@&PP&WPkYDG8)mDkKx~@!@*~rb9fr~l	N~dDDb^^/dS7s,@!@*,J`JPmx9~I;;+kY oKD:vE^A6Yrb~@!@*Pr5EPP4x@#@&,~P,P3p&KPji~@#@&P,3	N~&0@#@&@#@&P,j2U?(rg`E]:J#,xP"+;!ndYcsKDh`EPHwnJ*@#@&,Pj3U?q6HvJ]1rb,'~"+$E+kYcoWM:cJglhnr#@#@&@#@&,PsK.~2mm4PaP(U,In;!n/DRoGM:@#@&~,P~P&W,Vn6Yv6~8#,xPr^EPKD~a,'PkODG4b1OrW	PKD~6~x,JEP:4+	@#@&~,PP~3^/n@#@&~,P~,P,/YMKhw,'~I;;nkYRwGDs`6*@#@&P,P,P~P(W,kU/D.`!mCd`6bSrf)KAE*P@*,!,Ktx@#@&P,P~P,P~~,/YMP+swP{~/tm0fCYnckY.Khw*@#@&~,PP~~,2UN,(6@#@&,P,PP,PkODAxOkDX~x,6@#@&~P,PP,~G4NImGD[jY sbnV9/ca*P'~dDDP+s2@#@&~,P,P2	N,(0@#@&~Pg+aO@#@&@#@&~PK4L"n^WMNU+ORi29lO+@#@&P,W8%"+mG.9?nYc/^Wd@#@&@#@&Ax9~?!4~OaMAAA==^#~@%>

<%#@~^wQIAAA==j`A~`w9lY{&O+s/mKm4snvqYhKHw+B(O+ss^NSr.rT.CV!n~g+A#mVEnb@#@&@#@&,~,f(tPkYDU;^	tDn@#@&@#@&~,Pq6~V0Yv6.kT.mV;+Sq*'E@$r~K4+U@#@&PP~~,PdYMj$V	4+M+P{PsrNvr.kT.Cs!+~yb@#@&PP,3s/@#@&P~P~~,/ODU5Vqtn.P'~E,4+Mn,J~LP~Ek^Nq4+M+j;^`EPHw+rSqD+::z2+*@#@&P~P~~,/ODU5Vqtn.P'~dDDj;^	4+.PLPJ,bg9PrP'P~Ers9t.+U;Vv(O+ss^NSr.rT.CV!n#@#@&~~,2x[~&0@#@&@#@&,P~U+DPW(L"nmKD[?Y~x,?+M-+MRZMnCYr(LnmOcrb9rG$R"+^GMN/nOr#@#@&,~,W8NImWMNUnYcb^Yb\n/Kxx^YbWx,x~/DDG4/WUU@#@&~P,G4NIn^KDNjnDRjW!.1+~{PrjnGb:3PrP'PkY.(D+:kP4^P[,E~?AK,$/lOnTW.XY~',BE~LP1nAjlsE~LP~rB,JPLPkODU;s4+.n@#@&P,~W(LI^GD9?Y SG^0Kzw~',l[JKm3]nmN6x^z@#@&~,PKxPDMGD,Dn/!:n~	+6D@#@&,PPK8%ImKD[?nOcr2+	@#@&,P~jYPG8NInmK.9?nDP{P1KY4rxT@#@&@#@&2H9,?j~ns0AAA==^#~@%>

<%#@~^/AAAAA==j`A~"+DED	{:G{t+UE@#@&~~,q0,dYMbm1nd/d\^PxPEirPPtU@#@&P~~,PP]nkwGxkncIn9kM+mD`r4YDwlzJJ~',I+$;+kYRUn.\Djl.kC8^+d`rj2".3]|1b\3r#~[,EJ:CbxDzD09CYm{h+	E d4Y:^E#@#@&P,~3Vk+@#@&~P~~,P]+k2W	/n "+Nr.mO`r4DY2=zJJPLPkOD;WhwmxzuK:+hL#@#@&P,~3x9P&0@#@&3HGPjj~cEkAAA==^#~@%>








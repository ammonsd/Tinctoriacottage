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
<!--#INCLUDE FILE="../includes/members_tbl_setup.inc" -->
<!--#INCLUDE FILE="../includes/send_email.inc" -->
<!--#INCLUDE FILE="../includes/build_where_sql.inc" -->
<!--#INCLUDE FILE="../includes/logon_check.inc" -->
<!--#INCLUDE FILE="../includes/get_program_path.inc" -->
<!--#INCLUDE FILE="../includes/check_msg_variables.inc" -->
<!--#INCLUDE FILE="../includes/company_setup.inc" -->
<!--#INCLUDE FILE="../includes/display_edit_errors.inc" -->

<%#@~^ugMAAA==@#@&@#@&,P,Pf&H,dYMnL:dW^S,(~,(~,/YMPn:a~,q(~~r	Y3DMH4M@#@&@#@&PP~~kY.nThdW^,',M+DnThnmY4`rJb@#@&@#@&,~P,Zl^s~?Y!wmZGhalUX|(x6W@#@&,PP~^mVsPGCDl8m/{?Y!2`kY.ZK:2C	X1mh+*@#@&@#@&~P,P1lsV~JKoGx|/tmVcr+:Cr^{skkOcldaJB/YMZKhwmxz1m:nb@#@&@#@&~P,Pq6~]+$E/ORoGM: ZK;xDPx~ZPK4n	@#@&P,~,P~,ZmVV,fbdw^lz{wW.h@#@&P,~PAV/@#@&P,P,P~PrUD2.Dg8D,'~T@#@&P~~,P~PUnDPnMDt/okP{~ZM+CYr8%mYvE?1DkaOrxTRGk^YrG	l.Xrb@#@&P~~,PP~(6PcI5!+dDRwWDs`rjE(LE#,'~ErPl	[P"+;!ndYcsKDh`EKM+-kAJ*Px~rJ#~P4+U@#@&~,P~,P,PP,k	O2MDH4MPx~bxYA.Dg4D,Q~F@#@&,P~P~~,P~P.Dt/LdcbN[~bxO2M.g4.BPrH+k/mL+,?;4N+^O,I+$;kM+Nr@#@&P,P,P~P3U9P(0@#@&P,P~~,PqW~"+5EdDRoKDs`JtlbsK6OJ*Px~rJP:4+	@#@&,~~P,P,P~PrUD2.Dg8D,'~r	Y2..g4.P3~8@#@&,P,PP,P,~PD.Hkod zNN,rxD2DMH8DBPrHn/dCT+~KaY,In5!kDn[r@#@&P,~,P~,2	NP&0@#@&P,P~P,P(W,kxD3DM14M~@*PZP:tnx@#@&,P~P,~P,P~CMX2..KD\/Td,'~DMH/T/c(Y:d@#@&P~~,PP,~P,ml^s~fb/aVCXm3MDGDk@#@&,P~~,PP3sk+@#@&,~,P~,P,PP;l^sPhDGm/dmwWDs@#@&,PP,~~PAx9P(0@#@&,P~PAUN,qW@#@&@#@&w/kAAA==^#~@%>

<%#@~^EQAAAA==j`A~GkkwVmX|oWM:~8wUAAA==^#~@%>

   <html>
   <head>
   <meta name="robots" content="NOINDEX,NOFOLLOW">
   <title><%=#@~^DgAAAA==dDD/K:alxH1mh+sQUAAA==^#~@%> Broadcast e-Mail</title>
   </head>

   <body bgcolor="#008080" >

   <div align="CENTER">

   <table width="575" border="1" cellspacing="0" cellpadding="2" bordercolor="#000080">

      <tr>
        <td width="50%" align="CENTER" valign="TOP" bgcolor="#000095">
        <font face="Arial" size="4" color="#FFFFFF">

    	<b>Process e-Mail Distribution List</b></td>

      </tr>

      <tr>
        <td valign="TOP" bgcolor="#B4B4B4">
        <font face="Arial" size="2">
        <div align="CENTER">
        <table width="90" cellpadding="5">

        <form Action= <%=#@~^KQAAAA==E4YOa)JzJ,[,dYMnL:dW^~LPJhlbV{^rdYclkwEP/AwAAA==^#~@%>	 Method=POST>
        <%#@~^IQAAAA==oKD~Al1tPXPbUP"+5E/O 5E+Mz?DDk	LCwwAAA==^#~@%>
           <input type="Hidden" name="$<%=#@~^AgAAAA==a,mAAAAA==^#~@%>" value="<%=#@~^FgAAAA==];;/DRp!+Mz?DDrxT`abbQgAAA==^#~@%>">
        <%#@~^BQAAAA==H6O,vwEAAA==^#~@%>

   	    <font face="Arial" size="3">
      <tr>
    	  <td nowrap>
          <b>eMail Subject:</b></td>
          <td nowrap><input type="Text" name="Subj" size="50" maxlength="50""></td>
      </tr>
      <tr>
    	  <td>
          <b>eMail Text:</b>
          <font size='-2'>
          ENTER text to e-mail<br>or PASTE in text previously copied</font></td>
          <td nowrap><textarea name="MailText" cols="70" rows="22" wrap="YES"></textarea></td>
      </tr>
      <tr>
        <td colspan="2" nowrap>
        <hr></td>
      </tr>
      <tr>
        <td colspan="2" align="CENTER">
        <%#@~^IgAAAA==(6PjA?Uqrg`r)m1+d/d\sE*P',EjrPK4nUzQkAAA==^#~@%>
           <input type="Submit" value="Send e-Mail">
           &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <%#@~^BgAAAA==3	N~&05gEAAA==^#~@%>
        <input type="Submit" value="Preview without Send" name="Preview">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <input type="Reset" value="Clear">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <a href="http://<%=#@~^EAAAAA==dDD/K:alxHCKh+hocAYAAA==^#~@%>/registry_doc.html#BcEmail" target="Help_Win"><img src="../graphics/buttons/help.gif" width="12" height="16" border="0" align="top" alt="Visitor Registry Documentation"></a></td>
        </td>
      </tr>
     </table>
   	 </div>
     </td>
    </tr>
   </table>
   </div>
   </body>
   </html>

<%#@~^CQAAAA==~A19,?`APIQIAAA==^#~@%>


<%#@~^mg4AAA==j`A~hDKm+k/|oWM:@#@&@#@&~~,Pf&\PbxYUnUY;xDDSPrUD2.D;UYM~~r	Y2..g4.~,dDD3slbVHkoB~k	Y/EMrAUD~,dYM?;^	4+M+BPdY.\koSPbUYdW^@#@&@#@&~~,PdYM\ko~{P"+;!+kORwW.:vJ\CbVKaYr#@#@&@#@&P,P,kUYjn	Y/xD.P{PT@#@&PP~~bxO2M.;xOMP{P!@#@&@#@&P,P~/DDj5^t.+,'Pr~	CAIAPc$nRtlrV,JkkYD~{PXnd*J@#@&,~,PoKD,2l1t,aPbx~I;;nkYRwGDs@#@&,~~P,P,qWPsn6Yc6Bq#,'~EfJPP4x@#@&,~,P~,P,PPkYMj;^4+M+~x,/YMj;^t.nPLPrP)19~rP'P~;k^N	4D+j5^`hk9cX~+*~"+;!+kORwW.:v6bb@#@&P,~P,PPAU[P&0@#@&~P~~g+aY@#@&P,P~(6PIn5!+dYcoKDhvJhD+7kAJ*P@!@*,JE~mxN,rxkYDvj3?Uq}1cJi(GJb~r@$J*P@*T,KtnU@#@&~P,~,P~kYM?;^4nDPxPkY.j$V4nDP[,E~bgf,J~[~$!ksNq4+M+j5^`JnRtlrVrSU2jUq}1`rj&9J*#@#@&,P~~AxN,(0@#@&P,~~?Y,W8L]n1W.NUnY,'~jD\n.cZ.+mOr8N+1Y`rbG6f~R]+1W.[k+Yrb@#@&PP,~G4NImGD[jY b1Ok7+/G	x+^ObWUP{~kY.G4;Wx	@#@&~P,PG4NIn^KDNUnYc?W!.^+,',Jj2J3;K~e,oI}H~\AHA3]UJ~[,dDDj$Vqt+M+@#@&P,P~W(L]n1WD9j+DRranU@#@&@#@&P~P~(6PHWD~W(L]n1WD[jY 2}o,K4x@#@&P,P,~P,PG4NIn^KDNUnYcHW7nokM/D@#@&P~~,P~P,9W,4r^+PHGDPG4N]mGMNU+Yc2}o@#@&P~P,P~~,PP,^l^VPdGCN|H(Dd{9$|.CDbC4^+d@#@&PP~~,P~P,~,qW,Vx`DDbh`kY.n4WUng4D*bP{PFZ~Ptx@#@&~P~~,P~P,~P,P~dDDn4G	+H4M~{PEvJ,[P^+6O`kY.n4WUng4DBf#,[Prb~J,[,:rNcdDDKtKU+g4.SW~&b~LPEOr~LPhbNv/YMn4Gx18DBGS**@#@&,~P,PP,~~P,2	N~qW@#@&P~P,~P,P~~,/Y.3slrVtdTPx,Z4+m0HkL.mDrl(Vndv/YM\/T#@#@&~~P,P,P~P~~&0~I5E/O wWDhcrn.+7rhE*P@!@*PrJ,Ptx@#@&,P~~,PP,~P,PP,9GPStbVnPdOM2\lbsHko~@!@*PJE@#@&P~P,~,P~,P,PP,P,rxDSGm,'~r	/YMc/DD2tCrVs/T~^t.c8&b#@#@&P,P~~,PP~~,P~P,~b0~bxDSW1P{~!,Y4+	@#@&~,PP,~P,PP,~~P,P,P~kUOdW^P{~VxcdDD2\CbV\/Tb,_~8@#@&PP,P,~P,P~P,P~~,2x9~q6@#@&,~~P,P,P~P~~,P~P"n/aWUdR.rD+~VWD`dDDAHlbVtdoBkUYdW^~ PF*~[,J@!(.@*J@#@&,P~P~~,P~P,~P,P~dDD2\CbV\/T~{PhbNv/YM2tCk^HdoBkUOdWm,QP8#@#@&~~P,P,P~P~~,P~SKGw@#@&~~,PP~~,P~P,~,2p&K,fr@#@&,~P,P~P,P~~AV/(0,?2Uj(rg`rb^mndkS-VrbP{PEirPK4n	@#@&P,~,P~,P,PP,P,rxD2.Dg4.~{P?UNA:lbsc/DDUHPnjnM\nDBdYMKG)9NDS~kY.ZKhalUH1m:+BPkOD;EdYUD-ntlk^SP"+;!ndYcsKDh`Ej!4%J*SPkY.3slks\kob@#@&~,P~,P,PP,PAs/@#@&P,P~~,PP,~P,PPAp(K,f}@#@&P~~,P~P,~P,2U[,q0@#@&,P~P,~,P~,P&0PbxD3DM18D,'~T,KtU@#@&PP,~~P,P,P~P~~bxO?UY;xO.,'PrUD?nxD/	Y.,_,F@#@&P,~P,P~P,P~3^/+@#@&P,PP,~~P,P,P~PrUD2.D;UYMPx~bxY3.MZUYM~3Pq@#@&,PP,P,~P,P~2	N~(6@#@&,~P,PP,~~P,W(L]+^GMNj+D HK\nH6Y@#@&,P~P,~,PJKWa@#@&@#@&,~P,W8L"+^GMN?OR;VWkn@#@&,P,Pj+O~K4%I^WMNjnDP'~HKY4k	L@#@&~,P,2x9P&W@#@&@#@&P,P~(6PI5E/YcoGDs`rn.+-rhE#,@!@*,JE~:t+U@#@&P~P,~,PoKD,kxDSK^P{PqPDW~qZ@#@&,~P,PP,~~P"+kwGxdnc.kDnPr@!8.@*JP'~74/DdW@#@&~,P,PP,1aY@#@&~P,P~~,I+k2W	/+c	.kD+,J@!t*~mVro	xJrmnUD+DEE@*jd+,$MWAk+MPAmm0~A!YOW	POG,I+D;D	PYK~oWM:@!z4c@*E,[~\(/Dd0@#@&,PP~~,P3(&P,?i~@#@&PP,PAUN,qo@#@&@#@&~,PP&WPU2?U(61vJzm^+ddd\sJ*~@!@*PEirPK4n	@#@&P,~,P~,I/wKxknRqDrYPE@!(D@*@!8D@*@!4M@*@!tWPmVroUxrJ^+	O+MJE@*CePidD~qG~gWO,b!YtKDb"+9PWWMPP4b/Pw;x1YkKU~eC@!Jt*@*E~LP-4;.S6@#@&~,PP~~,2pq:~Uj$@#@&,PP,2	[P&0@#@&@#@&~~,PIdwKx/ 	DbYPE@!4OsV@*J,'P74/.d0@#@&~,P~IdaWUk+cDbY~J@!tnl9@*E~LP\(/Dd0@#@&~~P,I/2WUdR	DbO+,J@!ObYVn@*rP'PkOMZGswmxXglsnPLPEP~DGC9mlkOPtlk^~]+kE^Yd@!&ObYs+@*EPLP-8;DSW@#@&P~P,]/2Kxk+RqDbO+,J@!z4+C[@*JPL~\(ZDdW@#@&,P,P]+d2Kxd+c	DbYn~r@!4G[H@*EPL~74/MS6@#@&,P,~I/2W	/n qDkDnPr@!4M@*@!4M@*@!4.@*@!8M@*@!4M@*@!(D@*E,[P-8;DJ0@#@&,P~,I/wKxknRqDrYPE@!9k\,CVbox{EEZA1:2]JE@*rP'P78ZMSW@#@&PP~~"+dwKUk+ qDbY+,J@!Ol(Vn@*rP'~74ZMJ0@#@&P,~~I/aWU/n qDrY~J@!Y.@*@!YN@*@!(@*HEs8D~K0,?E1md/6EsPU+U[k)@!J8@*@!zY9@*EPLP74/DJW@#@&~P,~I/2G	/+ 	MkO+,E@!Y[@*@!(@*J,[,rxD?nxDZUOMP[,E@!J4@*@!&ON@*@!JY.@*E~LP-4;.S6@#@&~,PP]nkwGxknc.bYPJ@!YM@*@!DN@*@!(@*H;s4+M~W6Pj	RjE1m/d0;s,?nx9d)@!z8@*Lx4d2p[U4k2p@!&DN@*JPLP78ZMSW@#@&P~~,I+k2W	/+c	.kD+,J@!Y[@*@!4@*J,'PbxO3MDZUOMP'Pr@!J4@*@!zDN@*@!zD.@*rP'P74/.d0@#@&~P,PId2W	/R	DrOPE@!JOl(Vn@*rP[~-(Z.S6@#@&P~,P"+/aW	d+c.kD+~E@!zNb-@*rP[,-8ZMS6@#@&P~~,In/aGxk+ 	MkYn~r@!&4K[H@*E,[,\4;DdW@#@&P~P,IndaWxknRqDkDn~J@!z4YhV@*E,[~\(/Dd0@#@&@#@&2U[,?;4,oMgDAA==^#~@%>

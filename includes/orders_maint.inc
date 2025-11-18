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
<!--#INCLUDE FILE="../includes/check_date.inc" -->
<!--#INCLUDE FILE="../includes/company_setup.inc" -->
<!--#INCLUDE FILE="../includes/display_edit_errors.inc" -->
<!--#INCLUDE FILE="../includes/get_words.inc" -->
<!--#INCLUDE FILE="../includes/calculate_sales_tax.inc" -->
<!--#INCLUDE FILE="../includes/order_form_template.inc" -->
<!--#INCLUDE FILE="../includes/logon_check.inc" -->
<!--#INCLUDE FILE="../includes/build_where_sql.inc" -->
<!--#INCLUDE FILE="../includes/encrypt_decrypt.inc" -->
<!--#INCLUDE FILE="../includes/get_program_path.inc" -->
<!--#INCLUDE FILE="../includes/transfer_html.inc" -->

<%#@~^qygAAA==@#@&@#@&,P,Pf&H,rxD?;4BP(S,lDHj+^I+$dc#BPkY.In5}D[+MSP(Ws]msU[BP8W^\mkUD~,/YMnThSKm@#@&@#@&~~,PZmsV,?+D;2{;WswCxzm&xWW@#@&P,P~/mVV~9mYC4md{jY!w`kYM/WswCxH1Ch#@#@&~P,P/D.KosSKm~'~!YKosKlDtcEr#@#@&@#@&@#@&P,~,qW,I;E/D sKDh`rH)(gKIA}?r#P@!@*~JIJ,K4+U@#@&P~P,~P,qW~"+;;nkY p!nMXdDDbxovJ}HJ*P@!@*,JE~:t+	@#@&,PP,~~P,P,/OD]n$r.N.P{P]n$E+dOcp;+MzkY.bxT`J}1rb@#@&P~P,P~~,PP(GV`/+M]n;,',Y.En@#@&P~P,~P,2sdq0~];;+kOcsGM:vJrgJ*~@!@*PEJ,K4n	@#@&,~P,PP,~~PkYMIn;6.9+.P{~I;;nkYRoGM:cJ}Hr#@#@&P,PP,P,~P,4GV`/n."+;,xPDDE@#@&P,P,P~P3U9P(0@#@&P,P~3	NP(W@#@&@#@&,~,P(6P(WV`/.I;~',Y.;Pl	[P"+;!ndYcp!+.XdOMkUovE.r#~@!@*PJeE,K4+	@#@&P~,P,PPESKLW	PHWDP]n$EkMnN,YW,9r/aVmX~n.G1+d/[P}D[nMPsG.sPchJG,Z/,q	0W*@#@&~P,P3Vk+@#@&,PP,~P,ml^s~SKoKxmZ4n13cJK.NDdmslkUOcldwrSkY.;Wswl	XgC:#@#@&,P~~AxN,(0@#@&@#@&~~P,?Y~+..t/L/,xP;DnCD+r8%mO`rj1DraYbxocfb^YbWUlMXEb@#@&P,~P(WV"n2DbxDPxPO.!+@#@&@#@&P,P~(6PIn5!+dYcoKDhvJt+x!J*~@!@*PEJ,K4n	@#@&,~P,PP,^CV^P"+OE.U|KG{tnx!@#@&~,PP3U9P(0@#@&@#@&~,P,q0,I5E/ORwW.hvJHz(1:I25jE#,',JeJ~P4+U@#@&~P,P~~,nDG^/d{tCbxOxmxm{"n;!+dYk@#@&~,PPAs/q0,]n;!+kY p;nMXdYMrxT`Ej}J#~@!@*PEJ,P4+U@#@&,PP,P,~PE?4kawn[,rD9nD,KW9Cz@#@&P,P~P~~,/OD"n;}D[nMP'~];;+kOcp;DH/YMk	L`r?6J*@#@&~,PP,~P,4W^]nmwx9PxPi29lO+"nmKD[ckYD]n$r.N.Bf):2BJJ*@#@&~P,P3Vk+(W,I+$;+kYR5;nDH/DDrxLcrZ6J*~@!@*PEE,KtnU@#@&~P,~,P~,B;lx1+^s+9P6D9+.~:WNmz@#@&PP,~~P,PkY.In5}D[+M~',In5!+/O 5EnDHdDDr	ovJZ}J*@#@&,P~P,P~~(WV"nmwxN,x~jaNmYnIn^KD[`kOD"+56MN+.SrJSfzPA#@#@&P,PPAx9~q6@#@&@#@&P~~,q0,]+$E+kO p!+MXdY.r	ocJU6J*P@!@*,JJ~GMP]+$;/Ocp!+DH/D.k	ocJ;rEb,@!@*,EJ,KtU@#@&,P,P~P~~K4%I^WMNjnDRZsGk+@#@&,~,P~,P,?+DPK8L"+^WMNdnDP',HWDtk	L@#@&,P,P~P~~"+9ks~lMXjn^I+5dvFb@#@&~,P~,P,PlMXUnV"+5/vFb~{P/D.I;rM[nD@#@&,P~P~~,Pj2Ujq}1cEmDX]n$/E#,x,l.H?VI;k@#@&,P~P,P~~U2?U(rg`J:GOl^I;dJb~{Pq@#@&~P,P~~,PmCs^P	DbO{PMl	/0D|uKtSc/DDKLsSW1~[,JWM[nDk{slrxO m/2J,b@#@&P~~,2x[~&0@#@&@#@&,P~,q6P?A?U(rg`EKKYCs"+;kE#,@!@*,EEP,K4+U@#@&~,P~P,~Ifrh,lDzjV]+$dv?3U?&r1vJ:GYmV]+$/Eb*@#@&,~P,PP,dOD"+$r.Nn.{?3?U(rg`ECMXIn5kJb`8b@#@&~,P,PP,q6~?A?jq}1cE:WYmsI;/rb~@*,F,K4+U@#@&P~P,~P,P~~wWD~r	YjE(~{P+,YKP?A?U(rg`EKKYCs"+;kE#@#@&P,~~P,P,P~P~~mDz?sI;dcbxYj;(Oq#,x,?3U?&r1vJm.X"+5/r#cr	Y?!8#@#@&P,~~P,P,P~1naD@#@&P,~P,P~3	NP(W@#@&~P,~,P~U2U?q}1vEKKYCV"+5dr#P{~?A??&6H`rKKYCV]n$/E#,RP8@#@&~,PP~~,qWPU3U?(}1vJKKYmsI;dJ*P@!~8PK4nx@#@&P,~~P,P,P~?3jUq61vEKKYCs"+;dE*PxPrE@#@&~,P,PP,P,~?A?jq}1cEmDX"n;kJ#,x~Jr@#@&P~P~~,P3Vkn@#@&P~~,PP~~,Pj2Uj&rHvJmDX"+$dJ*PxPmDzjVI5/@#@&P,~~P,PAx[P(W@#@&~P,~2^/n(6P?3jUq61vE"+5}D9+Dr#,@!@*,JEP:tnU@#@&P,~P,PPkO.I;}D[+.~{Pj2Ujq}1cE"+;6.9+.J*@#@&P~,P,PPU2Ujq}1cJ"+56MN+ME#,'PrE@#@&,P,P3x[~&0@#@&@#@&P,P~(6PkdH!:nDb^v/OMI;rMN.#,'~0mVdn,WD,dYMI+$6.ND,'~JE~:tnx@#@&P,P~~,PmCs^P!+Dm}D[D|14M@#@&~P,P3Vk+@#@&,PP,~P,4W^]nmwx9PxP!nDr.N.fYCr^/`dOMIn;}.9+.*@#@&PP,P,~Pb0~4KV]n1sx9~',YD!n~K4+	@#@&P~~,P~P,~P(Ws]msU[,'~MO}D[D&Y+s/vdYMIn;}D[nM#@#@&~P,PP,~~P,k6P8Ws]mox9~',Y.;PK4n	@#@&P,~,P~,P,PP,P(GVhDGm/d6MN+M~',YD!n@#@&,P,P~P~~,P~P,^l^V~9b/wsCH{6D9nM{oKDs@#@&,P,~P,P~P,2sd@#@&,~P,PP,~~P,P,+.D\dT/ b9[P8~~EgWPJr	+~qDns/~wW!xN,0K.Pr[ErMNn.,aPr'/DDI56D9+M@#@&P~~,P~P,~PAx[~&0@#@&~,P~P,~AVd@#@&PP,P,~P,PnDMHdLkRb9[P8~Pr6.ND,a~J'dDD]+$6D9+.'rP1GO,sGE	[r@#@&,P,PP,PAUN,qW@#@&P~~,PP,(0,+DM\dokR;W;xO~@*PTP:4+	@#@&~,PP~~,P~Pm.H2.MWMH/T/,xPD.Hkod &Y+sd@#@&PP,~~P,P,P(0~];;+kOR5En.H/Y.r	ocJ}Hr#~@!@*,JJ,K4nx@#@&~P,P~~,PP,~P,/YM$Cm0Hko~'~EcJ@#@&,~P,P~~,PP3U9P(0@#@&,P~,P,PP,P1CV^P9kkwsCH{2M.WM/@#@&~~P,P,P3x[~&0@#@&,~P,2U[,q0@#@&@#@&ojg/:q6gPV+Y}D9nDG+OlbVdc}DN.1(D#@#@&@#@&,PV+Or.[D9+DCk^/~x,0lsd@#@&P,(6P6MND1(D,xPrJ~K4+U@#@&PP,~PA(q:~ojgZ:q61@#@&,P3x9~q6@#@&@#@&PPjnDPG4N]mGMNU+Y,',j+M\nDcZ.nmY+}8LmYvE)f}f~R]+^GMNd+DE#@#@&~~K4L]n1W.NUnDR6a+	P/DD}.ND9+DlrskK4^SPkYDG8/W	xBPCN62xjYmOk1~~C9SW^Vh+d/bhb/ObmBPl9Zs[Km4s+@#@&@#@&,Pq6~xKYPK8%ImKD[?nOc26s,Ptx@#@&,PP~~K4%I^KD[U+DRHK\okM/O@#@&P~~,PW(%ImWM[j+DRwkUN~crr.N.a,'~E,[P6.9+.1(.*@#@&,P,PP&0,UWDPG4NIn^KDNUnYc2rw~Ptx@#@&~P~~,P~PVnY}D[nMf+OCbVdP{~DD;@#@&PP,P,~P,/OD}D[nM14M~',W4N]nmKD9?nY ob+sNkcJ}D[nMaJb@#@&P~P,~,P~kYM2Hmk^~',W8L"+^GMN?ORwk+^[d`r+ HCksE*@#@&P,~P,P~~kYD6.9+.fmOPx,W(LImK.NU+ORwkns9/`r6D9+DGCO+r#@#@&~P~~,P~PkODglhn,'PG8NInmK.9?nDRwk+^NkcJglh+r#@#@&,PP,~P,P/D.jYM+Y~'~G(L]+1GD9?nOcskns9/cJUOM+nDJ*@#@&,P,~P,P~/DD)[9D+kd ,'PK8%ImKD[?nOcsr+^[/vJ)[9D+ddyJb@#@&~,P~,P,P/DD;rYHPxPK4%]mWM[?YRwrnV9/vJ/kOzr#@#@&,~P,P~~,/Y.jDlO+,x,W8NImWMNUnYcsr+^Ndcr?YmO+r#@#@&~~P,P,P~/O.\k2P{~W(L]n1WD[jY sbn^NdvJ\kwr#@#@&P,P~P,P~dDDZK;xDDX,x~W(L"+^W.[U+ORwr+^NdcrZW;UDDzJ*@#@&P~,P,PP,/D.n4WU+,'~G(LI^WMN?O sb+^Nd`EK4WU+g8Dr#@#@&,PP~~,P~/D.wla,',W4NI^WMNj+DRorVNkcJwl6g8.J*@#@&P~P~~,P~k	O?mVnd:l6~x,W8L"n1W.9?YRwksNk`E?mVnd:l6rb@#@&PP,~~P,PbxO?4rawrxT~',W8%"+mG.9?nYcob+s9/vJ?4ka2k	oE#@#@&~~,PP,~PbxY}.[+MKKYCV~x,W8L"nmKD[jYRorV[/vE}D[D:WYmVrb@#@&P~P,P~~,/YMjtbww[9lD+,'~W8%"+^WM[?Y ob+V[dvJjtb2a+[GlD+J*@#@&~P,P~P,PdOMZl	^+^flDn~',W(L]+^GMNj+D sb+s[k`J/C	mnVGCD+E*@#@&PP,P,~P,qWPDDrhvW4N]+1WD9jnYcsb+sNdcrZ/1(.J*#~@!@*PJE~:tnx@#@&,P~,P,PP,P,dYMZ/KHwn~{PW(%ImWM[j+DRwknV[dvJ/Z:zwJb@#@&PP~~,P~P,~,/OMZ;26afDnP{PG4NIn^KDNUnYcsks[/vJ;Z3629D+E#@#@&P,P~~,PP~~,/OD;/g4.,',f+;DH2YvW8L"+^GMN?ORwk+^[d`rZ;18DEbBS/lkn`K4%]mW.[U+ORwrV[k`r+OtlbsJ*#b@#@&P~~,PP,~2	NP&W@#@&,P,P~2U[,qW@#@&~P,P~G(LIn^KD[/OcZsK/@#@&,P,~PU+OPK4%]mWM[/YP{~HWDtbxL@#@&~,2UN,(0@#@&@#@&AxN~o`1/K&6g@#@&@#@&wj1;K&61,MnY}D[nMqYh/vrD9n.1(D*@#@&@#@&~,f(H,dYM?5sqt+.n@#@&@#@&,~V+O}D9+D&Yh/,'~0mVdn@#@&P,(0,rD9n.1(D,'~!~P4+U@#@&~P,P~3oqK~o`1/K&6g@#@&,PAxN,q6@#@&@#@&~PU+O~K4L"nmKDNUnOP{PU+.\n.cZ.+mO+}4%n1Y`E)Gr9Ac]mGMNk+Yr#@#@&P,W8L"+^GMN?ORzmYb-nZKx	+^YrG	PxPkODG4/G	x@#@&@#@&P~/D.U;sqtD+,',EPqtnDPE~LPA!rV9t.n?$VvJ(1 on{6D9nD[JS6MN+.H(Db@#@&~,W8NImWMNUnYc?GEMmn~{PJU3SAZK,,KDKN!mOZG[TSP]}EmxOrDXTS~]KrY^nY~~]nMkmTB~$&fDPwI6\,JPL~/DDrM[nD&Y:dK8s,[~J,EPLPdOM?;s	4+.+,',J~}IG2I,AI~$&fDJ@#@&@#@&,PW(%ImWM[j+DRdW^3Pza+~',CNdW^V"+l[6	Vz@#@&~,W8NImWMNUnYcr2+	@#@&@#@&PPbUYUE4,x~!@#@&@#@&~P(W,1GY,G4NIn^KDNjnDR3rw~:tn	@#@&PP,P,!+Dr.ND(O:/,xPDDE@#@&P,P,PG4%]mGD9j+DR\G7+sr.kY@#@&,~,P~GW,tbV~1KY~W(L]n1WD9j+DR2}o@#@&,P,P~P~~bxO?!8P{PrUD?E8~3Pq@#@&~,P~,P,PlMX}.ND(Y:dcbxYU;4B!#,x~W(L"+^W.[U+ORwr+^NdcrnDG[!mOZK[Jb@#@&,PP,P,~PmDzrMNn.&Y+sd`bxYU;8~8#,'~W8%"+^WM[?Y ob+V[dvJ}EmUDkOHJ*@#@&,P,~P,P~lMX6.9+D&O+s/`bUO?!4B bPx~K4%I^WMNjnDRsrn^Nd`rPbYsJ*@#@&,P,~P,P~lMX6.9+D&O+s/`bUO?!4B&bPx~K4%I^WMNjnDRsrn^Nd`rKMk^J*@#@&,P,~P,P~W(L]n1WD9j+DRHK-n16D@#@&P~~,PJWK2@#@&P~3	NP(W@#@&@#@&,~mDz}D9+D&Yh/v!S!*Px~bxYU;4@#@&P,G8L"+1W.NdnDR/VKd+@#@&~~U+Y~G(L]+1GMNjY,'PgWD4k	o@#@&@#@&3U9Ps`HZ:qrg@#@&@#@&?`A~n.G1+d/|\lbxOn	lx^n|In;!nkYd@#@&@#@&P,PG(H,/ODU;s	4+DSPkYDU5s4+M++~~8KVox9SPbxOj!4~~aBPdYM6MNnM1(D@#@&@#@&~P,Infb:~CMX?sI;/v]n;!+kY sG.sR/W!UY*@#@&~,PkUOUE8P{~Z@#@&,P,sWMPACm4PaP&x~];EdYcsWMh@#@&,P,P~BK.Kmn/k~l^V~9AS2P3,In;!nkYd,sbD/D@#@&~P,P~q6Psn6Y`XSF*P',E9J,K4+U@#@&~,P~P,~P&0~dDD?5sqtnD~{PErP:t+	@#@&~P,P~P,P~~,/YMj;^t.nP{PrP	tn.PE@#@&~P,P~~,PP~~kY.?$sqtnM+yP',J,	tDnPr@#@&~,PP,~P,2Vkn@#@&,P,P~P~~,P~/D.?$V	4D+~x,/ODU5^4DP[,J,6I,J@#@&,P~~,PP,~P,/YMj5VqtDn ~x,/ODU5Vqtn. P'~rP6I,E@#@&~,P,PP,PAUN,qW@#@&P~~,PP,~/DD?$s	tDPxPdOM?5Vq4+M+~',AErs94+MnU;svJ&1R}D9nD[JS:bNcaB #*@#@&,PP,~~P,/DDj;s	4+.+y~',/O.U;V	4Dn ,',A;bV9tDj;^`EqgRoF|rD9nD[J~sr[`X~y#b@#@&~,P~PAUN,qW@#@&PP~H6O@#@&~,P(6PkYDU;^	tDnP@!@*~ErPK4nx@#@&P,~~P,?Y~W8%"+^WM[?Y~x,?+.-D ZMnmYn}4N+mD`r)f}f$R"+^GMN/OJ*@#@&,~~P,PK4%In^KD[?ORzmOr7+ZGU	+^YbG	Px,/DDf(ZKUx@#@&~P,P~~K4L"nmKDNUnORUW!D^+~x,J92d3KAPo]}HPE~LPdYM6MNnMfYlbVkP4^P'PrPE~LP/D.?$V4n.+@#@&,P~P~~K4%I^WMNjnDRSG^0Kzw~{PC9SKm3"+m[r	Vz@#@&P~~,PPK8L"+mK.[?Ycr2+U@#@&P~P,~PU+O~K4L]n1W.NUnDPx,1KYtbxT@#@&,P~P,PjnDPW(%ImWM[j+DP{Pj+.-D ZMnlD+68N+mOcrb9rG$cIn1WMN/Yrb@#@&P~P,P~G(LI^WMN?O b1Yb\nZGU	+^YbGx,'~dDDf8/KxU@#@&~,P~,PK4L"+1GD9?nYc?G;Mm+,xPrf2d3P2,s"r\PE~LPdYM6D9+.(D+:dP(V~[,E,J~LPkYDU;^	tDn @#@&~~,PP,G4NI+1G.NU+DRJW^V:X2+,xPmNJG13InC9rUVH@#@&P~,P,PW(L"nmKD[?Y 6a+x@#@&P,PP,~j+DPK4%In^KD[?OP{PHGDtkUL@#@&~P,~,PjY,W4NI^WMNj+DPx~U+D7nDcZDCO+}4N+^YcEzf6f~ ImG.9/+OE*@#@&P,~,P~K4NI+1WM[?Y b1Yr-ZW	U+1YkKU~',/DD94/G	x@#@&,~P,P~G(LIn^KD[?Oc?G!D1+P{Pr92d2P2,s]6tPJ,'PkYD~rsVbxTfnYCr^/P4^~[,J~E,[PdOM?5Vq4Dny@#@&PP,P,~W(L]+1W.[U+YcJW13KH2nP{PmNJW^V"+CN}UVH@#@&~,PP~~K4%I^KD[U+DRra+	@#@&,P~P,PjnDPW(%ImWM[j+DP{PHWO4bxL@#@&~P,P~~(WV\CbxOP{~DD;@#@&PP,2	[P&0@#@&,P~8KVs	[P{P0msd+@#@&,P~sG.,2Cm4~6,qU~"+;;nkY sK.s@#@&,P,PPEnMGm/dPmVs~`nfzP2,I+$;n/D/@#@&~P~~,qWP^n0D`aS8#Px~rjEPmU9P];!+/DRwGDs`EfrP'~skNva~y##,x~JrP:tnx@#@&,P~P,~P,4GswxN~x,Y.E@#@&P~,P,PP,/D.rMNnDg4.~{P:b[`X~ *@#@&P,P,P~P~8KV]+1ox9Px~`wNCOInmK.9`dDD}DNDg8DBIn;!+dOcsWMh`r?fr'dYMrMNnDH8M#SI5E/O wWDhcrZ9JLdDD6MND1(D*b@#@&P~P,P3U9Pq6@#@&,PPgnaY@#@&,P~qW~(Wss	[P{PO.!+PO4x@#@&,~,P~,W(LImK.NU+OR;VGd@#@&,~P,PPUnOPK4NInmG.9/nY,xPgWO4bxo@#@&,P~2	[,qW@#@&,PPwWM~2mm4PXP(U,I+$;+kYRwG.:@#@&,P~P~vhDGmd/,lss,I2)9,YGPG(UnJz5,I+$EdYk@#@&P,P~~&0P^n0D`6BqbP{PrIEPCU9P]+$;+kY oKD:cEGJ~[,hbNcX~y##,',EJ,K4+	@#@&~,PP,~P,kxDj;4,',kUYj;(PQP8@#@&,P~~,PP~CMXj+^];dvk	Y?!4*~',:rNv6S+*@#@&,~P,P2	[~q6@#@&P~PHnXY@#@&,~P&0~r	Y?;8,@*~!,P4+U@#@&,PP,P,j2U?(rg`ECMXI5/r#P{~CDH?V]+5d@#@&~P,~P,?3jUqrHcrKGYms"+5kJ*P',k	O?!4@#@&,P~~,Pq6~4KVHmrUY,',Y.En~Dtnx@#@&P,P~~,PP^C^V~MrD+m:Dmx/6+MmC:HJ`kY.KT:SK^PLPJK.[+M/|:CkUOcldwr~#@#@&~~,PP~3	N~q6@#@&P~,2^/+@#@&,~P,P~BgW~iaNlDnP"+;!ndYkPtl[+@#@&,P~P,~mmVs~"+Y;.	{PW|\x;@#@&,PPAx9~q6@#@&@#@&2H9,?j~@#@&@#@&?`$~IY!DU{PG|Hnx!@#@&,P~(6P?3jUq61vEzm^/kS\^J*~',JiJ,K4n	@#@&,~P,PP"ndwKxk+ In[bDnmDcJ4YO2=zzE~LP]+$;/Oc?D\DjCDbl8V/cEU2Ij3I|1bt3E#,[,J&:Cr	Y&WM[+M/mhxE d4YhVrb@#@&~,PAV/@#@&~P,P~P"+d2Kx/ INkMn^YvJ4YOwl&JJ~[,dYMZGhalxzuK:nnTb@#@&~,PAxN,q6@#@&A19PUj$@#@&@#@&wi1;Kq}H~jaNmYnIn^KD[`}.NDH8M~?4rafCYS;lU1+^flD+*@#@&@#@&~P,4Gstlk	OP{PYM;n@#@&@#@&P~PjnDPG4N]+1W.[U+Y~x,?nD7nMR/M+mY+}4NnmD`EbGr9$cI+1GD9/+DEb@#@&P,PG4%]mGD9j+DR62xPdOMr.N.G+Omk^/K(VB~/DD94;WUUBPl96wx?DCOk1~,l[SG^0nn/kr:b/Or1~PC[;:[Km8^+@#@&@#@&PP,q6~W(L]+1W.[U+Yc3rwPK4nU@#@&P,P~P~~`w[lDnImG.9P'~WmVd+@#@&,P~AVk+@#@&P,~PK4%ImG.9?+D HK\+wr./D@#@&P~P~G(L]+1GD9?nOcskU[,`ErM[D:,',JPLP}.NDH4M#@#@&,PP,(0,W4N]nmKD9?nY 3}s~K4nx@#@&~~,PP~~`w[lDn"+^KD9P',0ms/@#@&P,P~3^/+@#@&P,PP,~~jaNmYnIn^KD[P{~YMEn@#@&PP~~,P~q6~;tn13GlY`;Cx1+sfmYnb,@!@*,EJ,KtU@#@&,P,P~P~~,PG4N]+1W.[U+Y ob+sNkcr?4bwa+NGlDnJ*PxPrJ@#@&,PP,~P,PP,G8L"+1W.NjnDRoksNk`E/mxmnsGlO+rb,'~;tm3GlDn`;lUmV9CD+#@#@&P,PP,~~2^/@#@&P~~,P~P,~PK4%]mW.[U+ORwrV[k`r?tbwanNGlO+r#~x,Zt^3GlYcjtbwGlO+b@#@&P~P,~P,P~~K4L]n1W.NUnDRob+^N/vJ;Cx1+sfmYnE*P',EJ@#@&P,~~P,PAx[P(W@#@&~P,~2	N~(6@#@&~~Ax[P&W@#@&@#@&P,W4NI^WMNj+DRi29lY@#@&@#@&2g9~s`1;K(rH@#@&@#@&?`$PV+Om}DNn.|18DZxoLAA==^#~@%>

<html>
<head>
<title>Display Order Request</title>
</head>
<body bgcolor="#6D6D6D">
<div align="CENTER">
<br><br><br><br><br><br>
<table border="1" cellspacing="0" cellpadding="2" bordercolor="#000080">
<tr>
<td align='CENTER' valign="TOP" bgcolor="#000095">
<font color="white">
Display Order Request
</font>
</td></tr>
<tr>
<td valign="TOP" bgcolor="#B4B4B4">
<font face="Arial" size="2">
<div align="CENTER">
<table border="0" cellspacing="0" cellpadding="5">
<form Action="orders_maint.asp" Method="POST">
<tr>
<td nowrap>
<b>Order Number:</b>
</td>
<td nowrap>
<input type="Text" name="ON" size="10" maxlength="10" value="">
</td>
</tr>
<tr>
<td nowrap colspan="2" align="center">
<input type="Submit" value="OK" Name="Submit">
</td>
</tr>
</form>
</table>
</div>
</td>
</tr>
</table>
</font>
</div>
</body>
</html>

<%#@~^CAAAAA==3gf~Uj~PAQIAAA==^#~@%>


<html><head>
<title>�A����</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<script language="JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

var LunData=new Array(
"0A4D0","0D250","1D295","0B550","056A0","0ADA2","095B0","14977","049B0","0A4B0",
"0B4B5","06A50","06D40","1AB54","02B60","09570","052F2","04970","06566","0D4A0",
"0EA50","16A95","05AD0","02B60","186E3","092E0","1C8D7","0C950","0D4A0","1D8A6",
"0B550","056A0","1A5B4","025D0","092D0","0D2B2","0A950","0B557","06CA0","0B550",
"15355","04DA0","0A5B0","14573","052B0","0A9A8","0E950","06AA0","0AEA6","0AB50",
"04B60","0AAE4","0A570","05260","0F263","0D950","05B57","056A0","096D0","04DD5",
"04AD0","0A4D0","0D4D4","0D250","0D558","0B540","0B6A0","195A6","095B0","049B0",
"0A974","0A4B0","0B27A","06A50","06D40","0AF46","0AB60","09570","04AF5","04970",
"064B0","074A3","0EA50","06B58","05AC0","0AB60","096D5","092E0","0C960","0D954",
"0D4A0","0DA50","07552","056A0","0ABB7","025D0","092D0","0CAB5","0A950","0B4A0",
"0BAA4","0AD50","055D9","04BA0","0A5B0","15176","052B0","06930","07954","06AA0",
"0AD50","05B52","04B60","0A6E6","0A4E0","0D260","0EA65","0D520","0DAA0","076A3",
"096D0","04AFB","04AD0","0A4D0","1D0B6","0D250","0D520","0DD45","0B5A0","056D0",
"055B2","049B0","0A577","0A4B0","0AA50","1B255","06D20","0ADA0","14B63","09370",
"049F8","04970","064B0","168A6","0EA50","06B20","1A6C4","0AAE0","092E0","0D2E3",
"0C960","0D557","0D4A0","0DA50","05D55","056A0","0A6D0","055D4","052D0","0A9B8");
var Today = new Date();
var Ny = Today.getFullYear();
var Nm = Today.getMonth();
var Nd = Today.getDate();
var cld,Selday;
var NHoliday = new Array(
"0101*����",
"0111 �q�k�`",
"0115 �Įv�`",
"0123 �ۥѤ�",
"0204 �A���`",
"0214 ���H�`",
"0215 ���@�`",
"0219 �s�ͬ��B�ʬ�����",
"0228*�M��������",
"0301 �L�и`",
"0305 ���l�x�`",
"0308 ���k�`",
"0312 �Ӿ�`",
"0317 ����`",
"0320 �l�F�`",
"0321 ��H�`",
"0325 ���N�`",
"0326 �s���`",
"0329 �C�~�`",
"0330 �X���`",
"0401 �M�H�`",
"0404 �����`",
"0405 ���ָ`",
"0407 �å͸`",
"0422 �@�ɦa�y��",
"0501*�Ұʸ`",
"0504 �����`",
"0505 �R�и`",
"0510 �]��`",
"0512 �@�h�`",
"0603 �T�ϸ`",
"0606 �u�{�v�`",
"0609 �K���`",
"0615 ĵ��`",
"0630 �|�p�v�`",
"0701 �����`",
"0711 ����`",
"0712 Ť�׸`",
"0808 ���˸`",
"0814 �ŭx�`",
"0827 �G���\�Ϩ�",
"0901 �O�̸`",
"0903 �x�H�`",
"0909 ��|�`",
"0913 �k�ߤ�",
"0928 �Юv�`",
"1006 �ѤH�`",
"1010*��y������",
"1021 �ع��`",
"1025 �x�W���_�`",
"1031 �U�t�`",
"1101 �ӤH�`",
"1111 �u�~�`",
"1117 �ۨӤ��`",
"1112 ����Ϩ���",
"1121 ���Ÿ`",
"1205 �����`",
"1210 �H�v�`",
"1212 �˧L�`",
"1225 ��ˬ�����",
"1227 �ؿv�v�`",
"1228 �q�H�`",
"1231 ���H�`");
var WHoliday = new Array(
"0520 ���˸`",
"0716 �X�@�`",
"0730 �Q���а�a�g",
"1144 �P���`");
var LHoliday = new Array(
"0101*�K�`",
"0102*�^�Q�a",
"0103*����",
"0104 �ﯫ",
"0105 �}��",
"0109 �Ѥ���",
"0115 ���d�`",
"0202 �Y��",
"0323 ������",
"0408 �D��`",
"0505*�ݤȸ`",
"0701 �}����",
"0707 �C�i���H�`",
"0715 �����`",
"0800 ������",
"0815*����`",
"0909 �����`",
"1208 þ�K�`",
"1216 ����",
"1224 �e��",
"0100*���i");

function initialize() {

 var Today = new Date();
var Ny = Today.getFullYear();
var Nm = Today.getMonth();
var Nd = Today.getDate();

 CLD.SY.selectedIndex=Ny-1912;
 CLD.SM.selectedIndex=Nm;
 drawCld(Ny,Nm);
}
var NMonthDays=new Array(31,28,31,30,31,30,31,31,30,31,30,31);
function isLeapYear(y,m) {
 if(m==1)
    return(((y%4 == 0) && (y%100 != 0) || (y%400 == 0))? 29: 28);
 else
    return(NMonthDays[m]);
}

function lYearDays(y) {
var i,k, sum = 0; 
  k=StrToInt(y,5);
  for(i=1;i<13;i++) sum += (k & (0x10000>>i))? 30 : 29;
   return(sum+leapDays(y));
}

function leapDays(y) {
  if(leapMonth(y)) return( (StrToInt(y,1)&0xf)? 30: 29);
  else return(0);
}

function leapMonth(y) {
 return(StrToInt(y,5) & 0xf);
}

function LmonthDays(y,m) {
  return( (StrToInt(y,5) & (0x10000>>m))? 30: 29 );
}

function StrToInt(yx,vx){
 var sr;
  sr=LunData[yx-1912];
   sr=sr.substring(0, vx);
   return (parseInt('0x'+sr));
}

function Lunar(objDate) {
 var i, leap=0, temp=0;
 var offset = (Date.UTC(objDate.getFullYear(),objDate.getMonth(),
                        objDate.getDate()) - Date.UTC(1912,1,18))/86400000;
 for(i=1912; i<2072 && offset>0; i++) { temp=lYearDays(i); offset-=temp; }
  if(offset<0) { offset+=temp; i--; }
   this.year = i;
  leap = leapMonth(i); 
  this.isLeap = false;
 for(i=1; i<13 && offset>0; i++) {
   if(leap>0 && i==(leap+1) && this.isLeap==false)
     { --i; this.isLeap = true; temp = leapDays(this.year); }
   else
    { temp = LmonthDays(this.year, i); }
     if(this.isLeap==true && i==(leap+1)) this.isLeap = false;
     offset -= temp;
 }

 if(offset==0 && leap>0 && i==leap+1)
    if(this.isLeap)
       { this.isLeap = false; }
    else
       { this.isLeap = true; --i; }
  if(offset<0){ offset += temp; --i; }
   this.month = i;
   this.day = offset + 1;
}
var Gan=new Array("��","�A","��","�B","��","�v","��","��","��","��");
var Zhi=new Array("�l","��","�G","�f","��","�x","��","��","��","��","��","��");
function GanZhi(num) {
 return(Gan[num%10]+Zhi[num%12]);
}


function calElement(sYear,sMonth,sDay,week,lYear,lMonth,lDay,isLeap,cYear,cMonth,jMonth,cDay) {
  this.isToday    = false;
  this.sYear      = sYear;   
   this.sMonth     = sMonth; 
    this.sDay       = sDay;  
     this.week       = week; 
  this.lYear      = lYear;   
   this.lMonth     = lMonth; 
    this.lDay       = lDay;  
     this.isLeap     = isLeap; 
  this.cYear      = cYear;   
   this.cMonth     = cMonth; 
    this.jMonth     = jMonth;
     this.cDay       = cDay; 
     this.color      = '';
  this.Lholiday = ''; 
   this.Nholiday = ''; 
    this.solarTerms    = '';
}
var TermData = new Array(0,21324,42505,63868,85407,107110,128977,151002,173218,195611,218134,240768,263418,286062,308631,331096,353423,375568,397546,419292,440895,462344,483626,504891);
function sTerm(y,n) {
 var offDate = new Date( ( 31556925974.7*(y-1912) + TermData[n]*60000  ) + Date.UTC(1912,0,7,0,7) );
 return(offDate.getUTCDate());
}
var SolarTerm = new Array("�p�H","�j�H","�߬K","�B��","���h","�K��","�M��","�\�B","�߮L","�p��","�~��","�L��","�p��","�j��","�߬�","�B��","���S","���","�H�S","����","�ߥV","�p��","�j��","�V��");
var dStr1 = new Array('��','�@','�G','�T','�|','��','��','�C','�K','�E','�Q');
function calendar(y,m) {
 var sDObj, lDObj, Ly, Lm, Ld=1, Lp, Lx=0, tmp1, tmp2, tmp3;
 var Cy, Cm, Jm, Cd; 
 var Ldpos = new Array(3);
 var n = 0;
 var firstLM = 0;
 sDObj = new Date(y,m,1,0,0,0,0); 
 this.length    = isLeapYear(y,m);
 this.firstWeek = sDObj.getDay(); 
  if(m<2)  Cy=GanZhi(y-1912+47); 
 else Cy=GanZhi(y-1912+48);
 var term2=sTerm(y,2);
 var firstNode = sTerm(y,m*2);
 Jm = GanZhi((y-1912)*12+m+36);
 var dayCyclical = Date.UTC(y,m,1,0,0,0,0)/86400000+21185+12;
 
 for(var i=0;i<this.length;i++) {
    if(Ld>Lx) {
       sDObj = new Date(y,m,i+1);
       lDObj = new Lunar(sDObj);
       Ly    = lDObj.year;
       Lm    = lDObj.month;
       Ld    = lDObj.day;
       Lp    = lDObj.isLeap;
       Lx    = Lp? leapDays(Ly): LmonthDays(Ly,Lm);
       if(n==0)  firstLM = Lm;
       Ldpos[n++] = i-Ld+1;
    }

    if(m==1 && (i+1)==term2) Cy=GanZhi(y-1912+48);
     if((i+1)==firstNode) Jm = GanZhi((y-1912)*12+m+37);
	Cm= GanZhi((Ly-1912)*12+Lm+37);
    Cd = GanZhi(dayCyclical+i);
    this[i] = new calElement(y,m+1,i+1,dStr1[(i+this.firstWeek)%7],
	                        Ly,Lm,Ld++,Lp,Cy,Cm,Jm,Cd );
 }

  tmp1=sTerm(y,m*2  )-1;
   tmp2=sTerm(y,m*2+1)-1;
 this[tmp1].solarTerms = SolarTerm[m*2];
  this[tmp2].solarTerms = SolarTerm[m*2+1];
   if(m==3) this[tmp1].color = 'red';

 for(i in NHoliday)
  if(NHoliday[i].match(/^(\d{2})(\d{2})([\s\*])(.+)$/))
   if(Number(RegExp.$1)==(m+1)) {
    if(Number(RegExp.$2)<=this.length){
     this[Number(RegExp.$2)-1].Nholiday += RegExp.$4 + ' ';
      if(RegExp.$3=='*') this[Number(RegExp.$2)-1].color = 'red';
       }
     }
 for(i in WHoliday)
  if(WHoliday[i].match(/^(\d{2})(\d)(\d)([\s\*])(.+)$/))
   if(Number(RegExp.$1)==(m+1)) {
    tmp1=Number(RegExp.$2);
     tmp2=Number(RegExp.$3);
      if(tmp1<5)
       this[((this.firstWeek>tmp2)?7:0) + 7*(tmp1-1) + tmp2 - this.firstWeek].Nholiday += RegExp.$5 + ' ';
      else {
       tmp1 -= 5;
        tmp3 = (this.firstWeek+this.length-1)%7;
         this[this.length - tmp3 - 7*tmp1 + tmp2 - (tmp2>tmp3?7:0) - 1 ].Nholiday += RegExp.$5 + ' ';
         }
       }
 for(i in LHoliday)
  if(LHoliday[i].match(/^(\d{2})(.{2})([\s\*])(.+)$/)) {
    tmp1=Number(RegExp.$1)-firstLM;
     if(tmp1==-11) tmp1=1;
      if(tmp1 >=0 && tmp1<n) {
       tmp2 = Ldpos[tmp1] + Number(RegExp.$2) -1;
        if( tmp2 >= 0 && tmp2<this.length && this[tmp2].isLeap!=true) {
         this[tmp2].Lholiday += RegExp.$4 + ' ';
          if(RegExp.$3=='*') this[tmp2].color = 'red';
         }
       }
    }
 if(m==2 || m==3) {
  var estDay = new easter(y);
   if(m == estDay.m)
    this[estDay.d-1].Nholiday = this[estDay.d-1].Nholiday+'�_���`';
 }
 if((this.firstWeek+12)%7==5) this[12].Nholiday += '�¦�P����';
 if(y==Ny && m==Nm) {
  this[Nd-1].isToday = true;
   this[Nd-1].color ='#ff00ff';
 }
}

function easter(y) {
 var term2=sTerm(y,5);
 var dayTerm2 = new Date(Date.UTC(y,2,term2,0,0,0,0)); 
 var lDayTerm2 = new Lunar(dayTerm2);

 if(lDayTerm2.day<15)
   var Lmlen= 15-lDayTerm2.day;
 else
   var Lmlen= (lDayTerm2.isLeap? leapDays(y): LmonthDays(y,lDayTerm2.month)) - lDayTerm2.day + 15;
 var NextMoon = new Date(dayTerm2.getTime() + 86400000*Lmlen );
  var dayEaster = new Date(NextMoon.getTime() + 86400000*( 7-NextMoon.getUTCDay()));
  this.m = dayEaster.getUTCMonth();
  this.d = dayEaster.getUTCDate();
}
var dStr2 = new Array('��','�Q','��','��','�m');
function cDay(d){
 var s;
 switch (d) {
    case 10:
       s = '��Q'; break;
    case 20:
       s = '�G�Q'; break;
    case 30:
       s = '�T�Q'; break;
    default :
       s = dStr2[Math.floor(d/10)];
       s += dStr1[d%10];
 }
 return(s);
}
var Zodiac=new Array("��","��","��","��","�s","�D","��","��","�U","��","��","��");
function drawCld(SY,SM) {
 var i,sD,s,size,Lastday;
 cld = new calendar(SY,SM);
 Lastday=isLeapYear(SY,SM);
  if (Selday> Lastday) Selday= Lastday;
 for(i=0;i<42;i++) {
    nObj=eval('SD'+ i);
    lObj=eval('LD'+ i);
	zObj=eval('DZ'+ i);
    nObj.style.cursor = 'hand'
    nObj.className = '';
   sD = i - cld.firstWeek;
    if(sD>-1 && sD<cld.length) {
      nObj.innerHTML = sD+1;
    if(cld[sD].isToday) Selday=sD+1;
    if (Selday==(i-cld.firstWeek+1)){
       mCls();
     if(cld[sD].lYear>1911) yDisplay = '����' + (((cld[sD].lYear-1911)==1)?'��':cld[sD].lYear-1911);
	   Ldz.innerHTML ='�@�A��:'+yDisplay+'�~ '+(cld[sD].isLeap?'�| ':' ')+cld[sD].lMonth+' �� '+cld[sD].lDay+' ��';
	 	Ldzx.innerHTML='�@�����G' +cld[sD].cYear+'�~ '+cld[sD].cMonth+'('+cld[sD].jMonth+')'+'�� '+cld[sD].cDay + '�� �ͨv�ݡG�i'+Zodiac[(cld[sD].lYear-4)%12]+'�j';
     if(cld[sD].sYear>1911) yDisplay = '����' + (((cld[sD].sYear-1911)==1)?'��':cld[sD].sYear-1911);
	  Ndz.innerHTML='�@���:'+yDisplay+'�~ '+cld[sD].sMonth+' �� '+cld[sD].sDay+' ��';
	   nObj.className = 'MoveColor';
	  } 
    nObj.style.color = cld[sD].color; 
     if(cld[sD].lDay==1) 
	  if (cld[sD].isLeap) 
	   lObj.innerHTML = '<b>'+'�|'+ cld[sD].lMonth + '��' + (leapDays(cld[sD].lYear)==29?'�p':'�j')+'</b>'
	  else
       lObj.innerHTML = '<b>'+ cld[sD].lMonth + '��' + (LmonthDays(cld[sD].lYear,cld[sD].lMonth)==29?'�p':'�j')+'</b>';
      else 
       lObj.innerHTML = cDay(cld[sD].lDay);
	   zObj.innerHTML = cld[sD].cDay;   
      s=cld[sD].Lholiday;
       if(s.length>0) {
        if(s.length>6) s = s.substr(0, 5);
         s = s.fontcolor('red');
       }
     else {
      s=cld[sD].Nholiday;
       if(s.length>0) {
        size = (s.charCodeAt(0)>0 && s.charCodeAt(0)<128)?8:5;
         if(s.length>size+2) s = s.substr(0, size);
          s=(s=='�¦�P����')?s.fontcolor('black'):s.fontcolor('blue');
        }
    else {
     s=cld[sD].solarTerms;
      if(s.length>0) s = s.fontcolor('#cc00cc');
       }
     }
    if(cld[sD].solarTerms=='�M��') s = '�M���`'.fontcolor('red');
     if(s.length>0) zObj.innerHTML = s;
    }
   else { 
    nObj.innerHTML = '';
    lObj.innerHTML = '';
	zObj.innerHTML = '';
    }
  }
}

function changeCld() {
 var y,m;
 y=CLD.SY.selectedIndex+1912;
 m=CLD.SM.selectedIndex;
 drawCld(y,m);
}

function pushBtm(K) {
 switch (K){
   case 'YD' :
    if(CLD.SY.selectedIndex>0) CLD.SY.selectedIndex--;
     break;
   case 'YU' :
    if(CLD.SY.selectedIndex<159) CLD.SY.selectedIndex++;
     break;
   case 'MD' :
    if(CLD.SM.selectedIndex>0) {
     CLD.SM.selectedIndex--;
    }
   else {
    CLD.SM.selectedIndex=11;
    if(CLD.SY.selectedIndex>0) CLD.SY.selectedIndex--;
     }
     break;
   case 'MU' :
    if(CLD.SM.selectedIndex<11) {
     CLD.SM.selectedIndex++;
     }
   else {
    CLD.SM.selectedIndex=0;
    if(CLD.SY.selectedIndex<159) CLD.SY.selectedIndex++;
     }
    break;
     default :
     CLD.SY.selectedIndex=Ny-1912;
     CLD.SM.selectedIndex=Nm;
 }
  changeCld();
}

 ////////////////////////////////////////
 var sk,st11,st22, cldt;
 function DATCld(SYY,SMM,TDD) {
 var i,sD,st,st1,st2,st3,size,Lastday;
   cldt = new calendar(SYY,SMM);
      //sD = i - cldt.firstWeek;
     sD = TDD ;
    if(sD>-1 && sD<cldt.length) {
    // if(cldt[SD].lYear>1911) yDisplay = '����' + (((cldt[SD].lYear-1911)==1)?'��':cldt[SD].lYear-1911);
	// Ldz.innerHTML ='�@�A��:'+yDisplay+'�~ '+(cldt[SD].isLeap?'�| ':' ')+cldt[SD].lMonth+' �� '+cldt[SD].lDay+' ��';
	// Ldzx.innerHTML='�@�����G' +cldt[SD].cYear+'�~ '+cldt[SD].cMonth+'('+cldt[SD].jMonth+')'+'�� '+cldt[SD].cDay + '�� �ͨv�ݡG�i'+Zodiac[(cldt[SD].lYear-4)%12]+'�j';
   //if(cldt[SD].sYear>1911) yDisplay = '����' + (((cldt[SD].sYear-1911)==1)?'��':cldt[SD].sYear-1911);
	//Ndz.innerHTML='�@���:'+yDisplay+'�~ '+cldt[SD].sMonth+' �� '+cldt[SD].sDay+' ��';
    // st1='����' + (((cldt[sD].lYear-1911)==1)?'��':cldt[sD].lYear-1911)+SYY+'�~ '+(SMM+1)+' �� '+(TDD+1)+' �� ';
	// st2 ='�@�A��:'+'�~ '+(cldt[sD].isLeap?'�| ':' ')+cldt[sD].lMonth+' �� '+cldt[sD].lDay+' ��';
	// st3 ='�@�����G' +cldt[sD].cYear+'�~ '+cldt[sD].cMonth+'('+cldt[sD].jMonth+')'+'�� '+cldt[sD].cDay + '�� �ͨv�ݡG�i'+Zodiac[(cldt[sD].lYear-4)%12]+'�j';
     st1 =cldt[sD].cMonth ;
     //st11=st1.substring(0,1);
     st11=st1.charAt(0)
     
     st2 =cldt[sD].cMonth ;
     st22=st2.charAt(1);
   //st1+st2+st3;
     // sk ='�@�����G'+cldt[sD].cYear ;
            }
    //
    return (st11+st22);
     
     //document.write (st11+st22);
    // Response.Redirect "Excer2-sn.asp?"& "YKK="& st11 & "&"& "YKG="& st22 & "&"& "LYG=" & st22 & "&" & "TNUM=" & (SMM+1) & "&" & "DNUM=" & (TDD+1) & "&" & "HNUM=" & st22
         
   //Response.Redirect "Excer2-sn.asp?" & "YKK=" & st1 & "&" & "YKG=" & st2 & "&" & "LYG=" & st2 & "&" & Request.QueryString
  // 
  
   }

//////////////////////////////////////

function mDown(v) {
 var s,ax;
 var nObj=eval('SD'+ v);
 var d=nObj.innerHTML-1;
 var  mObj=eval('SD'+ Selday);
  if(nObj.innerHTML!='') {
    mObj.className='';
	 mCls();
   if(cld[d].lYear>1911) yDisplay = '����' + (((cld[d].lYear-1911)==1)?'��':cld[d].lYear-1911);
	 Ldz.innerHTML ='�@�A��:'+yDisplay+'�~ '+(cld[d].isLeap?'�| ':' ')+cld[d].lMonth+' �� '+cld[d].lDay+' ��';
	 Ldzx.innerHTML='�@�����G' +cld[d].cYear+'�~ '+cld[d].cMonth+'('+cld[d].jMonth+')'+'�� '+cld[d].cDay + '�� �ͨv�ݡG�i'+Zodiac[(cld[d].lYear-4)%12]+'�j';
   if(cld[d].sYear>1911) yDisplay = '����' + (((cld[d].sYear-1911)==1)?'��':cld[d].sYear-1911);
	Ndz.innerHTML='�@���:'+yDisplay+'�~ '+cld[d].sMonth+' �� '+cld[d].sDay+' ��';
    nObj.className='MoveColor';
    Selday=v-cld.firstWeek+1;
     //  alert("are you sure ?");
     // Response.Redirect ("Excer2-1.asp?"& "YKK="&cld[d].cYear& "&" & "TNUM=" & cld[d].cMonth & "&" & "DNUM=" & cld[d].cDay & "&" & "HNUM= ��") 
    
    s='�@�����G' +cld[d].cYear+'�~ '+cld[d].cMonth+'('+cld[d].jMonth+')'+'�� '+cld[d].cDay + ' ��';

       } 
       
   
      
     }  
 
function GetZhi(s){
   s = s.substr(1, 1);
  for(i=0;i<12;i++)
   if(s==Zhi[i]) return(i);
}	

function mCls() {
 for(i=0;i<42;i++) {
    mObj=eval('SD'+ i);
	 mObj.className='';
  }	 
 }  
//-->
</script>
</head>
 <!--
<body bgcolor="#FFCC99" text="#000000" onload=initialize() onunLoad="MM_preloadImages('Images/News-1.gif','Images/FLoad-1.gif','Images/Product-1.gif','Images/Lun-1.gif','Images/GoodLink-1.gif','Images/GuestBook-1.gif')" background="Images/Hbar-7.gif">
 -->
 <body bgcolor="#FFCC99" text="#000000" onload=initialize() >
<div align="center"> 
  <table width="590" align="center" cellpadding="0" cellspacing="0" border="0">
    <tr>
     <!-- <td height="80" width="623" valign="top" align="center"> 
        <div align="center"><a href="http://syws.myweb.hinet.net/" target="_blank"><img src="Images/ShanYang-2.gif" width="560" height="70" border="0"></a></div>
      </td>
    </tr>
    <tr>
      <td height="22" width="590"> 
        <div align="center"><a href="News.htm" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image25','','Images/News-1.gif',1)"><img name="Image25" border="0" src="Images/News.gif" width="90" height="22"></a><img src="Images/Vbar.gif" width="4" height="22"><a href="downLoad.htm" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image17','','Images/FLoad-1.gif',1)"><img name="Image17" border="0" src="Images/FLoad.gif" width="90" height="22"></a><img src="Images/Vbar.gif" width="4" height="22"><a href="Proucts.htm" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image19','','Images/Product-1.gif',1)"><img name="Image19" border="0" src="Images/Product.gif" width="90" height="22"></a><img src="Images/Vbar.gif" width="4" height="22"><a href="LunCal.htm" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image20','','Images/Lun-1.gif',1)"><img name="Image20" border="0" src="Images/Lun.gif" width="90" height="22"></a><img src="Images/Vbar.gif" width="4" height="22"><a href="GoodLink.htm" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image21','','Images/GoodLink-1.gif',1)"><img name="Image21" border="0" src="Images/GoodLink.gif" width="90" height="22"></a><img src="Images/Vbar.gif" width="4" height="22"><a href="http://syws.gotdns.com/" target="_blank" onMouseOver="MM_swapImage('Image23','','Images/Home-1.gif',1)" onMouseOut="MM_swapImgRestore()"><img name="Image23" border="0" src="Images/Home.gif" width="90" height="22"></a></div>
      </td>-->
    </tr>
  </table>
  <table border=0 cellpadding=0 cellspacing=0 width=590 height="327">
    <tbody> 
    <tr> 
      <td align=right height=11 valign=bottom width="11"><img src="Images/Lun_Lup.gif" width="11" height="11"></td>
      <td bgcolor=#CCCC33 height=11 width="568"><font size=1>�@</font></td>
      <td align=left height=11 valign=bottom width="12"><img src="Images/Lun_Rup.gif" width="11" height="11"></td>
    </tr>
    <tr> 
      <td bgcolor=#CCCC33 width="11" height="309"><font size="1">�@</font></td>
      <td bgcolor=#ffffff width="580" height="309"> 
        <table width="568" height="304" bordercolor="#0000FF" bgcolor="#0000FF" border="0" align="center">
          <tr> 
            <td height="312" bgcolor="#CCCC33" valign="top" bordercolor="#CCCC33" background="Images/Bmp1.gif"> 
              <div align="center">
                <style>.MoveColor {
        			BACKGROUND-COLOR: aqua;
					}
					</style>
                <script language=JavaScript><!--
 if(navigator.appName == "Netscape" || parseInt(navigator.appVersion) < 4)
 document.write("<h1>�A���s�����L�k���榹�{���C</h1>���{���ݦb IE4 �H�᪺�����~�����!!")
//--></script>
              </div>
              <form name="CLD">
                <div align="center"> 
                  <table width="560" align="center" bordercolor="#CC9933" border="0" bgcolor="#CCCC99">
                    <tbody> 
                    <tr bgcolor="#CCCC33" align="center" bordercolor="#CC9900" valign="top"> 
                      <td colspan=7 height="18"><font color=#0000ff size=3 
            style="FONT-SIZE: 12pt"> </font> 
                        <table width="465" border="0" height="15" align="center" bgcolor="#CC9900">
                          <tr> 
                            <td align="center" bgcolor="#CCFFCC" height="15" width="460" valign="top" bordercolor="#333333"><font color=#0000ff size=3 
                        style="FONT-SIZE: 11pt"> 
                              <input type="button" name="Ybta"onClick="pushBtm('YD')" value="�ա�">
                              <select name=SY onChange=changeCld() 
                      style="FONT-SIZE: 11pt">
                                <script language=JavaScript><!--
           for(i=1912;i<2072;i++) 
 		   document.write('<option>'+i+'(����'+((i-1911)==1?'��':(i-1911))+')');
		            //--></script>
                              </select>
                              <font color=#0000ff size=3 style="FONT-SIZE: 12pt"> 
                              <input type="button" name="Ybtd" onClick="pushBtm('YU')" value="�ϡ�">
                              </font><font size="3" color="#0000FF">�~</font> 
                              <input type="button" name="Mbta"onClick="pushBtm('MD')" value="�ա�">
                              <select name=SM onChange=changeCld() 
            style="FONT-SIZE: 11pt">
                                <script language=JavaScript><!--
          for(i=1;i<13;i++) document.write('<option>'+i)
          //--></script>
                              </select>
                              <input type="button" name="Mbtd" onClick="pushBtm('MU')" value="�ϡ�">
                              <font color="#0000FF" size="3">��</font> 
                              <input type="button" name="Today" onClick="pushBtm('')" value="����">
                              </font></td>
                          </tr>
                        </table>
                        <div align="left"><FONT color=#000000 id=Ndz size=3></FONT> 
                          <FONT color=#ff0000 id=Ldz size=3></FONT> <br>
                          <FONT color=#ff00ff id=Ldzx size=3></FONT> </div>
                      </td>
                    </tr>
                    <tr align=center bgcolor=#99CC66 bordercolor="#CC9900"> 
                      <td width=80 height="19"><b><font color="#FF0000">�P����</font></b></td>
                      <td width=80 height="19"><b>�P���@</b></td>
                      <td width=80 height="19"><b>�P���G</b></td>
                      <td width=80 height="19"><b>�P���T</b></td>
                      <td width=80 height="19"><b>�P���|</b></td>
                      <td width=80 height="19"><b>�P����</b></td>
                      <td width=80 height="19"><b><font color="#FF0000">�P����</font></b></td>
                    </tr>
                    <script language=JavaScript><!--
          var gDay;
          for(i=0;i<6;i++) {
             document.write('<tr align=center>')
             for(j=0;j<7;j++) {
                gDay = i*7+j
				 document.write('<td id="GD' + gDay +'" onMouseDown="mDown(' + gDay +')"><font id="SD' + gDay +'" size=5 face="Arial Black"');
                if((j == 0) || (j == 6) ) document.write(' color=red');
              document.write(' TITLE=""></font><br><font id="LD' + gDay + '"size=3 style="font-size:10pt"> ')
        		document.write(' </font><br><font id="DZ' + gDay + '"size=3 style="font-size:10pt"> </font></td>')
		     }
             document.write('</tr>')
          }
          //--></script>
                    </tbody> 
                  </table>
                </div>
                <p align="center"><font color="#999900" face="�ө���">�ȴ����N���A����</font> 
                </p>
                
              </form>
            </td>
          </tr>
        </table>
      </td>
      <td align=middle bgcolor=#CCCC33 width="11" height="309"><font size="1">�@</font></td>
    </tr>
    <tr align=middle> 
      <td align=right height=11 valign=top width="11"><img src="Images/Lun_Ldown.gif" width="11" height="11"></td>
      <td bgcolor=#CCCC33 height=11 width="568"><font size=1>�@</font></td>
      <td align=middle height=11 valign=top width="12"><img border=0 height=11 
      src="Images/Lun_Rdown.gif" 
  width=11></td>
   </tr>
    </tbody> 
  </table>
</div>
</body>
</html>

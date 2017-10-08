<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>新增網頁1</title>
</head>
<body>
<%
  
  
 LunData =Array("0A4D0","0D250","1D295","0B550","056A0","0ADA2","095B0","14977","049B0","0A4B0","0B4B5","06A50","06D40","1AB54","02B60","09570","052F2","04970","06566","0D4A0")
 LunData2=Array("0EA50","16A95","05AD0","02B60","186E3","092E0","1C8D7","0C950","0D4A0","1D8A6","0B550","056A0","1A5B4","025D0","092D0","0D2B2","0A950","0B557","06CA0","0B550")
 LunData3=Array("15355","04DA0","0A5B0","14573","052B0","0A9A8","0E950","06AA0","0AEA6","0AB50","04B60","0AAE4","0A570","05260","0F263","0D950","05B57","056A0","096D0","04DD5")
 LunData4=Array("04AD0","0A4D0","0D4D4","0D250","0D558","0B540","0B6A0","195A6","095B0","049B0","0A974","0A4B0","0B27A","06A50","06D40","0AF46","0AB60","09570","04AF5","04970")
 LunData5=Array("064B0","074A3","0EA50","06B58","05AC0","0AB60","096D5","092E0","0C960","0D954","0D4A0","0DA50","07552","056A0","0ABB7","025D0","092D0","0CAB5","0A950","0B4A0")
Dim Today,Ny,Nm,Nd
 NToday = Now
 Ny = Year(NToday)
 Nm = Month(NToday)
 Nd = Day(NToday)
DIM cld,Seldayvar
 NHoliday1 = Array("0101*元旦","0111 司法節","0115 藥師節","0123 自由日","0204 農民節","0214 情人節","0215 戲劇節","0219 新生活運動紀念日","0228*和平紀念日","0301 兵役節","0305 童子軍節")
 NHoliday2 = Array("0308 婦女節","0312 植樹節","0317 國醫節","0320 郵政節","0321 氣象節","0325 美術節","0326 廣播節","0329 青年節","0330 出版節","0401 愚人節","0404 婦幼節","0405 音樂節")
 
 NHoliday3 = Array("0407 衛生節","0422 世界地球日","0501*勞動節","0504 文藝節","0505 舞蹈節","0510 珠算節","0512 護士節","0603 禁煙節","0606 工程師節","0609 鐵路節","0615 警察節","0630 會計師節")
 NHoliday4 = Array("0701 漁民節","0711 航海節","0712 聾啞節","0808 父親節","0814 空軍節","0827 鄭成功誕辰","0901 記者節","0903 軍人節","0909 體育節","0913 法律日","0928 教師節","1006 老人節","1010*國慶紀念日")
 WHoliday =  Array("0520 母親節","0716 合作節","0730 被奴役國家週","1144 感恩節")
 LHoliday1 = Array("0101*春節","0102*回娘家","0103*祭祖","0104 迎神","0105 開市","0109 天公生","0115 元宵節","0202 頭牙","0323 媽祖生","0408 浴佛節","0505*端午節","0701 開鬼門","0707 七夕情人節","0715 中元節")
 LHoliday2 = Array("0800 關鬼門","0815*中秋節","0909 重陽節","1208 臘八節","1216 尾牙","1224 送神","0100*除夕")

function initialize() 
 Dim Today,Ny,Nm,Nd
 Today = Date()
 Ny = Today.getFullYear()
 Nm = Today.getMonth()
 Nd = Today.getDate()

// CLD.SY.selectedIndex=Ny-1912
/// CLD.SM.selectedIndex=Nm
 END FUNCTION

NMonthDays=Array(31,28,31,30,31,30,31,31,30,31,30,31)
function isLeapYear(y,m) 
 if(m=1) then
    '''return(((y%4 == 0) && (y%100 != 0) || (y%400 == 0))? 29: 28)
   'IF ((y/4 = 0) and (y/100 <> 0) or (y/400 = 0)) THEN
    IF (((y Mod 4)=0) and ((y mod 100) <> 0)) or ((y mod 400) = 0)  THEN
        
      isLeapYear=29
    ELSE
      isLeapYear=28
    END IF
    
 else
    ''return(NMonthDays[m])
  isLeapYear=NMonthDays(m)
 end if   
 END FUNCTION
 
 MK =Array(&h10000,&h8000,&h4000,&h2000,&h1000,&h0800,&h0400,&h0200,&h0100,&h0080,&h0040,&h0020,&h0010,&h0008,&h0004,&h0002,&h0001)
function lYearDays(y) 
 Dim i,k
 sum = 0 
  ''hxy =("&h"&Mid(y,1,5))
  '' hxy =StrToInt(y,5)
  k=StrToInt(y,5)
   ''for(i=1;i<13;i++) sum += (k & (0x10000>>i))? 30 : 29;
   ''return(sum+leapDays(y));
   for i=1 to 12
    if (k and MK(i))<>0 then
        sum=sum+30
    else
        sum=sum+29
    end if        
   next
  lYearDays=sum+leapDays(y) 
   
END FUNCTION


function leapDays(y) 
  ''if(leapMonth(y)) return( (StrToInt(y,1)&0xf)? 30: 29)
  ''else return(0)
 if leapMonth(y)<>0 then
    ''hxy =("&h"&Mid(y,1,1))
    hxy =StrToInt(y,1)
    IF  (hxy and (&hf)) <>0 THEN 
     leapDays =30
    ELSE
      leapDays=29
    END IF
  else
    leapDays=0
  end if
  
END FUNCTION

function leapMonth(y) 
  ''return(StrToInt(y,5) & 0xf);
   ''hxy =("&h"&Mid(y,1,5))
  hxy =StrToInt(y,5)

 if (hxy and (&hf)) <>0 then
      leapMonth= (hxy and (&hf))
   else 
       leapMonth= 0
  end if

END FUNCTION


function LmonthDays(y,m) 
  ''return( (StrToInt(y,5) & (0x10000>>m))? 30: 29 );
    ''hxy =("&h"&Mid(y,1,5))
    hxy =StrToInt(y,5)
   if ( hxy and (&h10000)) >m then
       ''LDy=Hex(hxy and (&h10000))&"hellow; ok"
      ''tst=Hex((&had) and (&h88))
       LmonthDays=30
      else 
       ''LDy=Hex(hxy and (&h10000))&"hellow; no"
       LmonthDays=29
   end if
     '' LDy=hxy
   
END FUNCTION


function StrToInt(yx,vx)
 Dim sr,hsr
  sr=LunData(yx-1912)
   ''sr=LunData[yx-1912]
   '' sr=sr.substring(0, vx);
   '' return (parseInt('0x'+sr));
   ''sr=MID(1,VX)
  hsr =("&h"&Mid(sr,1,vx))
 StrToInt=hsr
END FUNCTION

 
 ''  Response.Write "<TR><TD>萬年干:"& ONLER & "</TD></TR>"
  ' Response.write  NToday&Ny&Nm&Nd 
 ''Response.write lYearDays(1925)&leapMonth(1925)
 '' Response.write  Lunar(1926,1,10)
 
 '''function Lunar(objDate)
 function Lunar(SECTM,TNUM,DNUM)
 Dim  i ,j,k,offsett,offsety,offsetm,offsetd,temp,tempy,tempm,tempd,offsettm
 Dim D1,D2,DY,DD
 leap=0
 temp=0
  D1=  #1912/2/18#
  D2=DateSerial(SECTM,TNUM,DNUM)
  DY=DateDiff("yyyy", D1, D2)
  DD=DateDiff("d", D1, D2) 
  ' YK8 = DY+8
  ' YK = YK8 MOD 10
 ' YG = DY MOD 12
 ' MG1= TNUM+1
  ' MG = MG1 MOD 12
  ' DK = DD MOD 10
  ' DG = DD MOD 12

''offset = (Date.UTC(objDate.getFullYear(),objDate.getMonth(),objDate.getDate())-Date.UTC(1912,1,18))/86400000)
  offsett = DateDiff("d", D1, D2) 
                      
 'for(i=1912; i<2072 && offset>0; i++) { temp=lYearDays(i); offset-=temp }
 ' if(offset<0) { offset+=temp; i--; }
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
  i=1912
 WHILE offsett >=0 
     temp=lYearDays(i)
      offsett= offsett-temp
      
   '' tempy=lYearDays(i)
    ''offsetm=offsett+temp
   i=i+1
  WEND 
   
  IF offsett<0 THEN
      i = i-1
     offsetm= offsett+temp
         
   END IF 
    
  thisyr = i
  leap = leapMonth(i)
  thisLeap = false
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ''j=0
   j=1
  WHILE (offsetm> 0 and j<13)
    ''j=j+1
     if(leap>0 and j=(leap+1) and thisLeap=false) then
      j=j-1
      thisLeap = true
      tempm = leapDays(thisyr) 
     else
      tempm = LmonthDays(thisyr, j)
     end if  
     if(thisLeap=true and j=(leap+1)) then
      thisLeap = false
     end if
    offsetm=offsetm-tempm
 
    'tempd=tempm
    ''offsetd=offsetm
   j=j+1
  WEND 

    IF(offsetm=0 and leap>0 and j=leap+1) THEN 
      if(thisLeap=true) then
        thisLeap = false 
      else
        thisLeap = true
        j=j-1
      end if 
    END IF
      IF (offsetm<0) THEN
        offsetd =offsetm+tempm
        j=j-1
      END IF  
   thismth = j
   thisday = offsetd + 1
  
  Lunars=SECTM&"年"&TNUM &"月"&DNUM&"日"
   ''  Lunar=offset&offsetm&offsetd
   '' Lunar=DD
  Lunar=Lunars&thisyr&"年"& thismth&"月"&thisday&"日總"&DD


 END FUNCTION
 
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Response.write lYearDays(1917)&leapMonth(1917)
  ' Response.write  DateDiff("d", D1, D2) 
  Response.write  Lunar(1917,12,31)
   '' Response.write isLeapYear(1916,2)
   ' ' Response.write ((1916 Mod 4)=0) 
      


 Gan=Array("甲","乙","丙","丁","戊","己","庚","辛","壬","癸")
 Zhi=Array("子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥")
Function GanZhi(num) 
 ' 'return(Gan[num%10]+Zhi[num%12])
END FUNCTION
%>

<script language=javascript>

 //var sDObjj = new Date(2012,6,1,0,0,0,0);
 var objDate= new Date(1917,11,31);

 var sDObjj = new Date(1912,1,18);
 var sDObjjy = sDObjj.getFullYear();
 var sDObjjm = sDObjj.getMonth();
 var sDObjjd = sDObjj.getDate();
 var offsetj = (Date.UTC(objDate.getFullYear(),objDate.getMonth(),objDate.getDate())-Date.UTC(1912,1,18))/86400000;
     
 // document.write(sDObjj);
  //document.write(sDObjjy);
  //document.write(sDObjjm);
 // document.write(Date.UTC(1912,1,18)/86400000);
  document.write('日總'+offsetj);
 
</script>
<script language=javascript>


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
  //END FUNCTION;

var TermData = new Array(0,21324,42505,63868,85407,107110,128977,151002,173218,195611,218134,240768,263418,286062,308631,331096,353423,375568,397546,419292,440895,462344,483626,504891);
function sTerm(y,n) {
 var offDate = new Date( ( 31556925974.7*(y-1912) + TermData[n]*60000  ) + Date.UTC(1912,0,7,0,7) );
 return(offDate.getUTCDate());
}
var SolarTerm = new Array("小寒","大寒","立春","雨水","驚蟄","春分","清明","穀雨","立夏","小滿","芒種","夏至","小暑","大暑","立秋","處暑","白露","秋分","寒露","霜降","立冬","小雪","大雪","冬至");
var dStr1 = new Array('日','一','二','三','四','五','六','七','八','九','十');

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
	                        Ly,Lm,Ld++,Lp,Cy,Cm,Jm,Cd);
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
    this[estDay.d-1].Nholiday = this[estDay.d-1].Nholiday+'復活節';
 }
 if((this.firstWeek+12)%7==5) this[12].Nholiday += '黑色星期五';
 if(y==Ny && m==Nm) {
  this[Nd-1].isToday = true;
   this[Nd-1].color ='#ff00ff';
 }
}
 // var sDObjj = new Date(2012,6,1,0,0,0,0);
 // document.write(sDObjj);


//////////////////////////////////////////////////////////
  
  var dStr2 = new Array('初','十','廿','卅','卌');
function cDay(d){
 var s;
 switch (d) {
    case 10:
       s = '初十'; break;
    case 20:
       s = '二十'; break;
    case 30:
       s = '三十'; break;
    default :
       s = dStr2[Math.floor(d/10)];
       s += dStr1[d%10];
 }
 return(s);
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

 </script>
<% 
  

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 function tst()
   
   if ((&had) and (&h88))<>0 then
      'tst="hellow; ok"
      tst=Hex((&had) and (&h88))
      else 
      tst="hellow; no"
   end if
   End function   
  
  function LDy(yt,m) 
    
    hxy =("&h"&Mid(yt,1,5))
    'hxy =(yt&m)
    if ( hxy and (&h10000)) >m then
       LDy=Hex(hxy and (&h10000))&"hellow; ok"
      ''tst=Hex((&had) and (&h88))
      else 
       LDy=Hex(hxy and (&h10000))&"hellow; no"
   end if
  '' LDy=hxy
   ''return( (StrToInt(y,5) & (0x10000>>m))? 30: 29 );('0x'+sr)
  end function

 '' Response.Write "----------LYGG------"
 '' Response.Write tst()
 ''  Response.Write LDy("dacdf",6)
   %>

</body>

</html>

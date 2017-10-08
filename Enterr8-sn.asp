<HTML>
<HEAD><TITLE>表單與 ASP 程式的合體</TITLE></HEAD>
<BODY BGCOLOR=#FFFFFF>
<H2 ALIGN=CENTER>表單與 ASP 程式的合體<HR></H2>
<% If Request("Send") = Empty Then %>
   <FORM Action=Forma-1.asp Method=POST>
   <INPUT Type=Hidden Name=Send Value="Send">
     <!--- 國曆年份：<INPUT Type= Text Name=Syerr Size=20 value=2012><P>
       月份：<INPUT Type=Text Name=Smohh Size=20 value=6><P>
       日期：<INPUT Type=Text Name=Sdatt Size=20 value=26><P>
       <!---<form action="<%=Myself%>" method="GET"> 
        <p>科目：<select name="Lesson" size="1"> 
            <option IsSelected("ZEWU", Lesson)>ZEWU</option> 
        </select></p> 
        <p>請輸入任意數字：<input type="text" size="20" name="No" Value="<%=No%>"></p>     
          姓名：<input type="text" size="20" name="Name"  Value="<%=Name%>"> 
        <p>性別：<select name="Sex" size="1">                                                                                                                                                                                                                                                                                                                 
            <option value="男">男</option><option value="女">女</option> 
             </select>
          </p>--> 

           
      <h5> 西元?年 　 &月份              &日號            &時辰                      &今年西元?年 :</h5>                                                                                                                                                                                                               
         <p>：<select name="YNUM" size="6"> 
            <option value="1912">1912</option><option value="1913">1913</option><option value="1914">1914</option><option value="1915">1915</option> 
            <option value="1916">1916</option><option value="1917">1917</option><option value="1918">1918</option><option value="1919">1919</option> 
             <option value="1920">1920</option><option value="1921">1921</option><option value="1922">1922</option><option value="1923">1923</option>
             <option value="1924">1924</option><option value="1925">1925</option><option value="1926">1926</option><option value="1927">1927</option>
             <option value="1928">1928</option><option value="1929">1929</option><option value="1930">1930</option><option value="1931">1931</option>
             <option value="1932">1932</option><option value="1933">1933</option><option value="1934">1934</option><option value="1935">1935</option>
             <option value="1936">1936</option><option value="1937">1937</option><option value="1938">1938</option><option value="1939">1939</option>
             <option value="1940">1940</option><option value="1941">1941</option><option value="1942">1942</option><option value="1943">1943</option>
             <option value="1944">1944</option><option value="1945">1945</option><option value="1946">1946</option><option value="1947">1947</option>
             <option value="1948">1948</option><option value="1949">1949</option><option value="1950">1950</option><option value="1951">1951</option>
             <option value="1952">1952</option><option value="1953">1953</option><option value="1954">1954</option><option value="1955">1955</option>
             <option value="1956">1956</option><option value="1957">1957</option><option value="1958">1958</option><option value="1959">1959</option>
             <option value="1960">1960</option><option value="1961">1961</option><option value="1962">1962</option><option value="1963">1963</option>
             <option value="1964">1964</option><option value="1965">1965</option><option value="1966">1966</option><option value="1967">1967</option>
             <option value="1968">1968</option><option value="1969">1969</option><option value="1970">1970</option><option value="1971">1971</option>
             <option value="1972">1972</option><option value="1973">1973</option><option value="1974">1974</option><option value="1975">1975</option>
             <option value="1976">1976</option><option value="1977">1977</option><option value="1978">1978</option><option value="1979">1979</option>
             <option value="1980">1980</option><option value="1981">1981</option><option value="1982">1982</option><option value="1983">1983</option>
             <option value="1984">1984</option><option value="1985">1985</option><option value="1986">1986</option><option value="1987">1987</option>
             <option value="1988">1988</option><option value="1989">1989</option><option value="1990">1990</option><option value="1991">1991</option>
             <option value="1992">1992</option><option value="1993">1993</option><option value="1994">1994</option><option value="1995">1995</option>
             <option value="1996">1996</option><option value="1997">1997</option><option value="1998">1998</option><option value="1999">1999</option>
             <option value="2000">2000</option><option value="2001">2001</option><option value="2002">2002</option><option value="2003">2003</option>
             <option value="2004">2004</option><option value="2005">2005</option><option value="2006">2006</option><option value="2007">2007</option>
             <option value="2008">2008</option><option value="2009">2009</option><option value="2010">2010</option><option value="2011">2011</option>
             <option value="2012">2012</option><option value="2013">2013</option><option value="2014">2014</option><option value="2015">2015</option>
                      
              </select> 
           <!.../p...>                                                                                                                                                                                                                                                                                                                                                                  
          ：<select name="MNUM" size="6">                                                                                                                                                                                                                                                                                                                                                                  
            <option value="1">1</option><option value="2">2</option> 
             <option value="3">3</option><option value="4">4</option> 
             <option value="5">5</option><option value="6">6</option> 
             <option value="7">7</option><option value="8">8</option> 
             <option value="9">9</option><option value="10">10</option> 
             <option value="11">11</option><option value="12">12</option> 
            </select>
            <!.../p...>                                                                                                                                                                                                               
          ：<select name="DNUM" size="6">                                                                                                                                                                                                                                                                                                                                                                    
             <option value="1">1</option><option value="2">2</option> 
             <option value="3">3</option><option value="4">4</option> 
             <option value="5">5</option><option value="6">6</option> 
             <option value="7">7</option><option value="8">8</option> 
             <option value="9">9</option><option value="10">10</option> 
             <option value="11">11</option><option value="12">12</option> 
             <option value="13">13</option><option value="14">14</option> 
             <option value="15">15</option><option value="16">16</option> 
             <option value="17">17</option><option value="18">18</option> 
             <option value="19">19</option><option value="20">20</option> 
             <option value="21">21</option><option value="22">22</option> 
             <option value="23">23</option><option value="24">24</option> 
             <option value="25">25</option><option value="26">26</option> 
             <option value="27">27</option><option value="28">28</option> 
             <option value="29">29</option><option value="30">30</option> 
             <option value="31">31</option> 
            </select>
            <!.../p...>                                                                                                                                                                                                               
          ：<select name="HNUM" size="6">                                                                                                                                                                                                                                                                                                                                                                    
            <option value="子">23~01</option><option value="丑">01~03</option> 
             <option value="寅">03~05</option><option value="卯">05~07</option> 
             <option value="辰">07~09</option><option value="已">09~11</option> 
             <option value="午">11~13</option><option value="未">13~15</option> 
             <option value="申">15~17</option><option value="酉">17~19</option> 
             <option value="戌">19~21</option><option value="亥">21~23</option> 
            </select>
                                                                                                                                                                               
         ：<select name="LYR" size="6">                           
             <option value="2007">2007</option><option value="2008">2008</option>  <option value="2009">2009</option><option value="2010">2010</option>
              <option value="2011">2011</option><option value="2012">2012</option><option value="2013">2013</option><option value="2014">2014</option>
             <option value="2015">2015</option><option value="2016">2016</option><option value="20176">2017</option><option value="2018">2018</option>
             <option value="2019">2019</option>  <option value="2020">2020</option><option value="2021">2021</option>
                <!.../p...>     
            </select>                            
          </p> 
        
   <INPUT Type=Submit Value="傳 送">
   </FORM>
<% Else %>
   <Center><H2>
   <!--<%=Request("Syerr")%> <%=Request("Smohh")%> <%=Request("Sdatt")%>-->
 
 <%
   YNUM=Request("YNUM")
   MNUM=Request("MNUM")
   DNUM=Request("DNUM")
   
   Dim Syern,Smohn,Sdatn
    Syern=Cint(YNUM)
    Smohn=Cint(MNUM)
    Sdatn=Cint(DNUM)
   ''  Slunar=Lunar(Syern,Smohn,Sdatn)
  '' Slunar=Lunar(1915,3,6) '''''''''''須在于後呼叫'''''''''  
 ''  Response.write  SLunar
'' ReDim LunarRY(31,12)  
 Dim LunDataT(200) 
 LunData =Array("0A4D0","0D250","1D295","0B550","056A0","0ADA2","095B0","14977","049B0","0A4B0","0B4B5","06A50","06D40","1AB54","02B60","09570","052F2","04970","06566","0D4A0")
 LunData1 =Array("0A4D0","0D250","1D295","0B550","056A0","0ADA2","095B0","14977","049B0","0A4B0","0B4B5","06A50","06D40","1AB54","02B60","09570","052F2","04970","06566","0D4A0")
 LunData2=Array("0EA50","16A95","05AD0","02B60","186E3","092E0","1C8D7","0C950","0D4A0","1D8A6","0B550","056A0","1A5B4","025D0","092D0","0D2B2","0A950","0B557","06CA0","0B550")
 LunData3=Array("15355","04DA0","0A5B0","14573","052B0","0A9A8","0E950","06AA0","0AEA6","0AB50","04B60","0AAE4","0A570","05260","0F263","0D950","05B57","056A0","096D0","04DD5")
 LunData4=Array("04AD0","0A4D0","0D4D4","0D250","0D558","0B540","0B6A0","195A6","095B0","049B0","0A974","0A4B0","0B27A","06A50","06D40","0AF46","0AB60","09570","04AF5","04970")
 LunData5=Array("064B0","074A3","0EA50","06B58","05AC0","0AB60","096D5","092E0","0C960","0D954","0D4A0","0DA50","07552","056A0","0ABB7","025D0","092D0","0CAB5","0A950","0B4A0")
 '''LunDataT=LunData+LunData2+LunData3+LunData4
 LunData6=Array("0BAA4","0AD50","055D9","04BA0","0A5B0","15176","052B0","06930","07954","06AA0","0AD50","05B52","04B60","0A6E6","0A4E0","0D260","0EA65","0D520","0DAA0","076A3")

LunData7=Array("096D0","04AFB","04AD0","0A4D0","1D0B6","0D250","0D520","0DD45","0B5A0","056D0","055B2","049B0","0A577","0A4B0","0AA50","1B255","06D20","0ADA0","14B63","09370")

LunData8=Array("049F8","04970","064B0","168A6","0EA50","06B20","1A6C4","0AAE0","092E0","0D2E3","0C960","0D557","0D4A0","0DA50","05D55","056A0","0A6D0","055D4","052D0","0A9B8")

 
 for k=0 to 19
  LunDataT(k)=LunData1(k)
  LunDataT(20+k)=LunData2(k)
  LunDataT(40+k)=LunData3(k)
  LunDataT(60+k)=LunData4(k)
  LunDataT(80+k)=LunData5(k)
  LunDataT(100+k)=LunData6(k)
  LunDataT(120+k)=LunData7(k)
  LunDataT(140+k)=LunData8(k)
 next
 Dim Today,Ny,Nm,Nd
 NToday = Now
 Ny = Year(NToday)
 Nm = Month(NToday)
 Nd = Day(NToday)
  'D11=  #1912/2/18#
  'D22=DateSerial(SECTM,TNUM,DNUM)
  'DYY=DateDiff("yyyy", D1, D2)
  'DDD=DateDiff("d", D1, D2) 

 
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
 
 MK =Array(&h10000,&h08000,&h04000,&h02000,&h01000,&h00800,&h00400,&h00200,&h00100,&h00080,&h00040,&h00020,&h00010,&h00008,&h00004,&h00002,&h00001)
function lYearDays(y) 
 Dim i,k
 sum = 0 
  ''hxy =("&h"&Mid(y,1,5))
  '' hxy =StrToInt(y,5)
  ''k=StrToInt(y,5)
  
  k= StrToInt2(y,4)
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

 Dim hxym ,hxymk
function LmonthDays(y,m) 
  ''return( (StrToInt(y,5) & (0x10000>>m))? 30: 29 );
     '''hxy =("&h"&Mid(y,1,5))  ''(k and MK(i))<>0
     ''hxy =StrToInt(y,5)
      hxym= StrToInt2(y,4)
    ''hxymk =Hex(hxym and MK(m))
   if  (hxym and MK(m))>m then
       ''LDy=Hex(hxy and (&h10000))&"hellow; ok"
       ''tst=Hex((&had) and (&h88))
       LmhDays=30
     else 
       ''LDy=Hex(hxy and (&h10000))&"hellow; no"
       LmhDays=29
   end if
     '' LDy=hxy
     LmonthDays=LmhDays
END FUNCTION
  
 '' Response.write  LunDataT(159)&"<BR>"
'' Response.write    LmonthDays(1919,1)&"日"&hxym&"<BR>"
 ''Response.write    (Hex((&h4977)and (&h08000)))&"<BR>"
 
Dim overm1 ,ovnerm2,overm11 ,overm22
 overm1 = Hex((&h14977)and (&h08000))
 overm2 =     (&h14977)and (&h08000)
 overm11 =Mid(overm1,2,4)
 overm22 =Mid(overm2,2,4)
 
'' Response.write overm1&"<BR>"
 ''Response.write overm11
 ''Response.write overm2&"<BR>"
'' Response.write overm22

function StrToInt(yx,vx)
 Dim sr,hsr
  sr=LunDataT(Abs(yx-1912))
   ''sr=LunData[yx-1912]
   '' sr=sr.substring(0, vx);
   '' return (parseInt('0x'+sr));
   ''sr=MID(1,VX)
  hsr =("&h"&Mid(sr,1,vx))
 StrToInt=hsr
END FUNCTION
function StrToInt2(yx,vx)
 Dim sr2,hsr2
  sr2=LunDataT(Abs(yx-1912))
   ''''sr=LunData[yx-1912]
   ''''' sr=sr.substring(0, vx);
   ''''' return (parseInt('0x'+sr));
   ''sr=MID(1,VX)
  hsr2 =("&h"&Mid(sr2,2,vx))
 StrToInt2=hsr2
END FUNCTION

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
        'j=j-1
      end if 
       j=j-1
    END IF
      IF (offsetm<0) THEN
        offsetd =offsetm+tempm
        j=j-1
      END IF
        
   thismth = j
   thisday = offsetd + 1
  
 '' Lunars=SECTM&"年;"&TNUM &"月;"&DNUM&"日;"
   ''  Lunar=offset&offsetm&offsetd
   '' Lunar=DD
 '' Lunar=Lunars&thisyr&"年"& thismth&"月"&thisday&"整總日"&DD
  ''LunarRY(DNUM,5)=thisyr
  
  Lunar= thisyr&";"& thismth&";"&thisday&";"&thisLeap
  
  
 END FUNCTION
 
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' Response.write lYearDays(1917)&leapMonth(1917)
  ' Response.write  DateDiff("d", D1, D2) 
' Response.write  Lunar(1917,12,31)
   '' Response.write isLeapYear(1912,2)
   ' ' Response.write ((1916 Mod 4)=0) 
   ''  Slunar=Lunar(1915,3,6)
   ''Response.write     hxy and MK(m)
 ' '' Response.write ("年總日"&lYearDays(1912)&"月總日"&lMonthDays(1912,4)&"<BR>") 
  Dim SLunar,SPLunar
  Slunar=Lunar(Syern,Smohn,Sdatn)
    SPLunar=split(Slunar,";")
   '' Response.write SLunar& SPLunar(3)

    Response.write SLunar
  TermDatav =Array(0,21324,42505,63868,85407,107110,128977,151002,173218,195611,218134,240768,263418,286062,308631,331096,353423,375568,397546,419292,440895,462344,483626,504891)
 Function ssTerm(y,nn) 
  Dim DDA , DDD 
  D1t=  #1912/1/7#
  '' D2t=DateSerial(2012,2,4)
  ''DYY=DateDiff("yyyy", D1, D2t)
  ''DDt=DateDiff("n", D1t, D2t)
 Mn =( 525948.77*(y-1912) + TermDatav(nn) ) 
 DDA=DateAdd("n",Mn,D1t)
 DDD= Day(DDA)
 ssTerm=DDD
End Function 

 ''Dim ssTermt
 '' ssTermt=ssTerm(2012,5)
  ''  Response.write ssTermt
   '' Response.write (DDt)&"<BR>"


  
 Gan=Array("甲","乙","丙","丁","戊","己","庚","辛","壬","癸")
 Zhi=Array("子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥")
Function GanZhi(num) 
 ' 'return(Gan[num%10]+Zhi[num%12])
Dim Gant ,Zhit
 Gant=Gan(num mod 10)
 Zhit=Zhi(num mod 12)
 GanZhi=Gant+Zhit
END FUNCTION

%>

<script language=javascript>

 //var sDObjj = new Date(2012,6,1,0,0,0,0);
 var objDate= new Date(1912,11,31);
  //var sDObjjd = sDObjj.getDate();
 var offsetj = (Date.UTC(objDate.getFullYear(),objDate.getMonth(),objDate.getDate())-Date.UTC(1912,1,18))/86400000;
  /// function calendar(y,m) {
  ///  this[i] = new calElement(y,m+1,i+1,dStr1[(i+this.firstWeek)%7],
  ///	                       Ly,Lm,Ld++,Lp,Cy,Cm,Jm,Cd); }
 
 var TermData = new Array(0,21324,42505,63868,85407,107110,128977,151002,173218,195611,218134,240768,263418,286062,308631,331096,353423,375568,397546,419292,440895,462344,483626,504891);
function sTerm(y,n) {
 
 var offDate = new Date( ( 31556925974.7*(y-1912) + TermData[n]*60000  ) + Date.UTC(1912,0,7,0,7) );
 return(offDate.getUTCDate());
 }
   sTermt=sTerm(2012,5) ;
    var offDateDN=( 31556925974.7*(2012-1912) + TermData[2]*60000  ) /60000 ;
  // document.write('測年節日'+sTermt);
   //document.write(Date.UTC(1912,0,7,0,7));
  // document.write('測年節分距'+offDateDN);
 
 </script>
 <%
 
  SolarTerm = Array("小寒","大寒","立春","雨水","驚蟄","春分","清明","穀雨","立夏","小滿","芒種","夏至","小暑","大暑","立秋","處暑","白露","秋分","寒露","霜降","立冬","小雪","大雪","冬至")
 dStr1 = Array("日","一","二","三","四","五","六","七","八","九","十")
 
 Function Tcalendar(y,m,dx)
 Dim sDO2, sDOj,lDOj, Ly, Lm,Lp,Lpp,thfsw ,tmp1, tmp2, tmp3,firstNode
  Ld=0
  Lx=1
 Dim Cy, Cm, Jm, Cd 
 Dim Ldpos(3),Tcalary(12,33,12)
  n = 0
 firstLM = 0
 '' Dim DDA , DDD 
 ' D1t=  #1912/1/7#
 ' '' D2t=DateSerial(2012,2,4)
  '''DYY=DateDiff("yyyy", D1, D2t)
  '''DDt=DateDiff("n", D1t, D2t)
'' Mn =( 525948.77*(y-1912) + TermDatav(nn) ) 
'' DDA=DateAdd("n",Mn,D1t)
 ''DDD= Day(DDA)

 
 Dim DCA , DCD,DCLD 
  D1D=  #1912/1/7# 
  D2D= #1912/1/1#
  sDO2=DateSerial(y,m,1)
   ''sDO2=sDObj = new Date(y,m,1,0,0,0,0) 
  ''DCY=DateDiff("yyyy", D1D, sDO2)
  ''DCN=DateDiff("n", D1D, sDO2)
   CM =( 525948.77*(y-1912) + TermDatav(nn) ) 
   DCA=DateAdd("n",CM,D1D)
   DCD= Day(DCA)
  thisfstW=DatePart("w",sDO2)-1
  thisln=isLeapYear(y,m-1)
  
  if((m-1)<2) then
    Cy=GanZhi(y-1912+47) 
  else
   Cy=GanZhi(y-1912+48)
  end if
  term2=ssTerm(y,2)  ''''立春''''
  firstNode = ssTerm(y,(m-1)*2)   ''''當月的首節氣為幾日''''
   Cm = GanZhi((y-1912)*12+(m-1)+36)

  ''Jm = GanZhi((y-1912)*12+(m-1)+36)
   ''var dayCyclical = Date.UTC(y,m,1,0,0,0,0)/86400000+21185+12
  DCLD=DateDiff("d", D2D, sDO2)+12
   '' for(var i=0;i<this.length;i++) {
  for i=1 to thisln 
   '' if(Ld<Lx) then
       ''sDObj = new Date(y,m,i+1)
       ''lDObj = new Lunar(sDObj)
      sDOj=DateSerial(y,m,i)
      Syerp=Year(sDOj)
      Smohp=Month(sDOj)
      Sdatp=Day(sDOj)
      thfsw=DatePart("w",sDOj)-1
     lDOj= Lunar(Syerp,Smohp,Sdatp) 
       lDOJV=Split(lDOj,";")
       Ly    = lDOJV(0)
       Lm    = lDOJV(1)
       Ld    = lDOJV(2)
       Lp    = lDOJV(3)
        if (Lp= "True") then
          Lx=leapDays(Ly)
         else
          Lx=LmonthDays(Ly,Lm)
         end if     
       if(n=0)  then
        firstLM = Lm
         ''Ldpos[n++] = i-Ld+1
        end if
    '' end if
        
    ' if(m==1 && (i+1)==term2) Cy=GanZhi(y-1912+48);
    '  if((i+1)==firstNode) Jm = GanZhi((y-1912)*12+(m-1)+37);
	'Cm= GanZhi((Ly-1912)*12+Lm+37);
    '  Cd = GanZhi(dayCyclical+i);
    
    if((m-1)=1 and i=term2) then
       Cy=GanZhi(y-1912+48)
     end if
  if(i=firstNode) then
     ''Jm = GanZhi((y-1912)*12+(m-1)+37)
    Cm= GanZhi((Ly-1912)*12+(m-1)+37)
 end if
	'' Cm= GanZhi((Ly-1912)*12+Lm+37)
     Cd = GanZhi(DCLD+i-1)
    
    ''' this[i] = new calElement(y,m+1,i+1,dStr1[(i+this.firstWeek)%7],
	'''                        Ly,Lm,Ld++,Lp,Cy,Cm,Jm,Cd);
        Tcalary(m,i,0)=y
        Tcalary(m,i,1)=m
        Tcalary(m,i,2)=i
        Tcalary(m,i,3)= thfsw
        Tcalary(m,i,4)=Ly
        Tcalary(m,i,5)=Lm
        Tcalary(m,i,6)=Ld
        Tcalary(m,i,7)=Lp
        Tcalary(m,i,8)=Cy
        Tcalary(m,i,9)=Cm
        Tcalary(m,i,10)=Cd
       '' Tcalary(m,i,11)=Jm
 next 
 
  '' tmp1=sTerm(y,m*2  )-1;
  '' tmp2=sTerm(y,m*2+1)-1;
 ''this[tmp1].solarTerms = SolarTerm[m*2];
 '' this[tmp2].solarTerms = SolarTerm[m*2+1];
 ''  if(m==3) this[tmp1].color = 'red';

  ''Tcalendar="<日期>"&D2D&"<星期>"&DCW&"<天數>"&DCL
  ''Tcalendar=Ly&"<年期>"&Lm&"<月期>"&Ld&"<日期>"&Lp&"<任否日>"
   ''Tcalendar="<國月>"&thisln&"<農月>"&Ld&lDOJV(2)&"<日期>"&Lx&Lp&Lpp
    '' Tcalendar="<測試期>"&Tcalary(m,21,6)&"<測農期>"&Tcalary(m,21,4)&Tcalary(m,21,5)&Tcalary(m,21,6)&"<全月日數>"&thisln&"<BR>"
      Tcalendar1="<測國期>"&Tcalary(m,dx,0)&"/"&Tcalary(m,dx,1)&"/"&Tcalary(m,dx,2)&"<星期>"&Tcalary(m,dx,3)&"<BR>"
     Tcalendar2= "<測農期>"&Tcalary(m,dx,4)&"/"&Tcalary(m,dx,5)&"/"&Tcalary(m,dx,6)&"<年月日干支>"&Tcalary(m,dx,7)&Tcalary(m,dx,8)&Tcalary(m,dx,9)&Tcalary(m,dx,10) &"<BR>"
    
    Tcalendar=Tcalendar1&Tcalendar2&"<測日期>"&Ld
 End function
 
 Dim sTcalendar
     'sTcalendar=Tcalendar(2012,12)
    sTcalendar=Tcalendar(Syern,Smohn,Sdatn)
    Response.write  sTcalendar
     Response.write  sDO2&"<BR>"
 
 %>
 <script language=javascript>
 
 
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

 

 End function
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

 <HR></H2></Center>
<% End If %>
</BODY>
</HTML>

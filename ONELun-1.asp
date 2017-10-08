<!--
        ***************************************
         農曆月曆&世界時間 DHTML 程式 (台灣版)
        ***************************************
                最後修改: 2003/8/12


如果您覺得這個程式不錯，您可以自由轉寄給親朋好友分享。自由使
用範圍: 學校、學會、公會、公司內部、程式研究、個人網站供人查
詢使用。

Open Source 不代表放棄著作權，任何形式之引用或轉載前請來信告
知。如需於「商業或營利」目的中使用此部份之程式碼或資料，需取
得本人書面授權。

最新版本與更新資訊於 http://sean.wox.org/ap/calendar/ 公佈


　                            歡迎來信互相討論研究與指正誤謬
                      連絡方式：http://sean.o4u.com/contact/
                      　　　　　　　　　　Sean Lin (林洵賢)
                          尊重他人創作‧請勿刪除或變更此說明

-->

<HTML><HEAD><TITLE>萬年曆</TITLE>
<META content="時間;世界時間;時區;農曆;陰曆;陽曆;月曆;曆法;節日;紀念日;時區;節氣;二十四節氣;干支;生肖;world time clock;gregorian solar;chinese lunar;calendar" name=keywords>
<META content=All name=robots>
<META content="Gregorian Solar Calendar and Chinese Lunar Calendar" name=description>
<META http-equiv=Content-Type content="text/html; charset=big5">
<SCRIPT language=JScript>
<!--
/*****************************************************************************
                                 個人偏好設定
*****************************************************************************/

var conWeekend = 3;  // 星期六顏色顯示: 1=黑色, 2=綠色, 3=紅色, 4=隔週休


/*****************************************************************************
                                 日期資料
*****************************************************************************/

var lunarInfo=new Array(
0x4bd8,0x4ae0,0xa570,0x54d5,0xd260,0xd950,0x5554,0x56af,0x9ad0,0x55d2,
0x4ae0,0xa5b6,0xa4d0,0xd250,0xd295,0xb54f,0xd6a0,0xada2,0x95b0,0x4977,
0x497f,0xa4b0,0xb4b5,0x6a50,0x6d40,0xab54,0x2b6f,0x9570,0x52f2,0x4970,
0x6566,0xd4a0,0xea50,0x6a95,0x5adf,0x2b60,0x86e3,0x92ef,0xc8d7,0xc95f,
0xd4a0,0xd8a6,0xb55f,0x56a0,0xa5b4,0x25df,0x92d0,0xd2b2,0xa950,0xb557,
0x6ca0,0xb550,0x5355,0x4daf,0xa5b0,0x4573,0x52bf,0xa9a8,0xe950,0x6aa0,
0xaea6,0xab50,0x4b60,0xaae4,0xa570,0x5260,0xf263,0xd950,0x5b57,0x56a0,
0x96d0,0x4dd5,0x4ad0,0xa4d0,0xd4d4,0xd250,0xd558,0xb540,0xb6a0,0x95a6,
0x95bf,0x49b0,0xa974,0xa4b0,0xb27a,0x6a50,0x6d40,0xaf46,0xab60,0x9570,
0x4af5,0x4970,0x64b0,0x74a3,0xea50,0x6b58,0x5ac0,0xab60,0x96d5,0x92e0,
0xc960,0xd954,0xd4a0,0xda50,0x7552,0x56a0,0xabb7,0x25d0,0x92d0,0xcab5,
0xa950,0xb4a0,0xbaa4,0xad50,0x55d9,0x4ba0,0xa5b0,0x5176,0x52bf,0xa930,
0x7954,0x6aa0,0xad50,0x5b52,0x4b60,0xa6e6,0xa4e0,0xd260,0xea65,0xd530,
0x5aa0,0x76a3,0x96d0,0x4afb,0x4ad0,0xa4d0,0xd0b6,0xd25f,0xd520,0xdd45,
0xb5a0,0x56d0,0x55b2,0x49b0,0xa577,0xa4b0,0xaa50,0xb255,0x6d2f,0xada0,
0x4b63,0x937f,0x49f8,0x4970,0x64b0,0x68a6,0xea5f,0x6b20,0xa6c4,0xaaef,
0x92e0,0xd2e3,0xc960,0xd557,0xd4a0,0xda50,0x5d55,0x56a0,0xa6d0,0x55d4,
0x52d0,0xa9b8,0xa950,0xb4a0,0xb6a6,0xad50,0x55a0,0xaba4,0xa5b0,0x52b0,
0xb273,0x6930,0x7337,0x6aa0,0xad50,0x4b55,0x4b6f,0xa570,0x54e4,0xd260,
0xe968,0xd520,0xdaa0,0x6aa6,0x56df,0x4ae0,0xa9d4,0xa4d0,0xd150,0xf252,
0xd520);

var solarMonth=new Array(31,28,31,30,31,30,31,31,30,31,30,31);
var Gan=new Array("甲","乙","丙","丁","戊","己","庚","辛","壬","癸");
var Zhi=new Array("子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥");
var Animals=new Array("鼠","牛","虎","兔","龍","蛇","馬","羊","猴","雞","狗","豬");
var solarTerm = new Array("小寒","大寒","立春","雨水","驚蟄","春分","清明","穀雨","立夏","小滿","芒種","夏至","小暑","大暑","立秋","處暑","白露","秋分","寒露","霜降","立冬","小雪","大雪","冬至");
var sTermInfo = new Array(0,21208,42467,63836,85337,107014,128867,150921,173149,195551,218072,240693,263343,285989,308563,331033,353350,375494,397447,419210,440795,462224,483532,504758);
var nStr1 = new Array('日','一','二','三','四','五','六','七','八','九','十');
var nStr2 = new Array('初','十','廿','卅','卌');
var monthName = new Array("JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC");

//國曆節日 *表示放假日
var sFtv = new Array(
"0101*元旦 中華民國開國紀念日",
"0111 司法節",
"0115 藥師節",
"0123 自由日",
"0204 農民節",
"0214 情人節",
"0215 戲劇節",
"0219 新生活運動紀念日",
"0228*和平紀念日",
"0229 測試",
"0301 兵役節",
"0305 童子軍節",
"0308 婦女節",
"0312 植樹節 國父逝世紀念日",
"0317 國醫節",
"0320 郵政節",
"0321 氣象節",
"0325 美術節",
"0326 廣播節",
"0329 青年節 革命先烈紀念日",
"0330 出版節",
"0401 愚人節 主計節",
"0404 婦幼節",
"0405 音樂節",
"0407 衛生節",
"0422 世界地球日",
"0501*勞動節",
"0504 文藝節",
"0505 舞蹈節",
"0510 珠算節",
"0512 護士節",
"0603 禁煙節",
"0606 工程師節 水利節",
"0609 鐵路節",
"0615 警察節",
"0630 會計師節",
"0701 漁民節 公路節 稅務節",
"0706 合作節",
"0711 航海節",
"0712 聾啞節",
"0808 父親節",
"0814 空軍節",
"0827 鄭成功\誕辰",
"0901 記者節",
"0903 軍人節 抗戰紀念",
"0909 體育節 律師節",
"0913 法律日",
"0928 教師節 孔子誕辰",
"1006 老人節",
"1010*國慶紀念日",
"1021 華僑節",
"1025 台灣光復節",
"1031 萬聖節 蔣公誕辰紀念日 榮民節",
"1101 商人節",
"1111 工業節 地政節",
"1117 自來水節",
"1112 國父誕辰紀念日 醫師節 中華文化復興節",
"1121 防空節",
"1205 海員節 盲人節",
"1210 人權節",
"1212 憲兵節",
"1225 行憲紀念日 民族復興節 聖誕節",
"1227 建築師節",
"1228 電信節",
"1231 受信節");

//某月的第幾個星期幾。 5,6,7,8 表示到數第 1,2,3,4 個星期幾
var wFtv = new Array(
"0520 母親節",
"0716 合作節",
"0730 被奴役國家週",
"1144 感恩節");

//農曆節日
var lFtv = new Array(
"0101*春節",
"0102*回娘家",
"0103*祭祖",
"0104 迎神",
"0105 開市",
"0109 天公生",
"0115 元宵節 觀光節",
"0202 頭牙 土地公生",
"0323 媽祖生",
"0408 浴彿節",
"0505*端午節 詩人節",
"0701 開鬼門",
"0707 七夕情人節",
"0715 中元節",
"0800 關鬼門",
"0815*中秋節",
"0909 重陽節",
"1208 臘八節",
"1216 尾牙",
"1224 送神",
"0100*除夕");

//世界時間資料


/*****************************************************************************
                                    日期計算
*****************************************************************************/

//====================================== 傳回農曆 y年的總天數
function lYearDays(y) {
 var i, sum = 348;
 for(i=0x8000; i>0x8; i>>=1) sum += (lunarInfo[y-1900] & i)? 1: 0;
 return(sum+leapDays(y));
}

//====================================== 傳回農曆 y年閏月的天數
function leapDays(y) {
 if(leapMonth(y)) return( (lunarInfo[y-1899]&0xf)==0xf? 30: 29);
 else return(0);
}

//====================================== 傳回農曆 y年閏哪個月 1-12 , 沒閏傳回 0
function leapMonth(y) {
 var lm = lunarInfo[y-1900] & 0xf;
 return(lm==0xf?0:lm);
}

//====================================== 傳回農曆 y年m月的總天數
function monthDays(y,m) {
 return( (lunarInfo[y-1900] & (0x10000>>m))? 30: 29 );
}


//====================================== 算出農曆, 傳入日期物件, 傳回農曆日期物件
//                                       該物件屬性有 .year .month .day .isLeap
function Lunar(objDate) {

 var i, leap=0, temp=0;
 var offset   = (Date.UTC(objDate.getFullYear(),objDate.getMonth(),objDate.getDate()) - Date.UTC(1900,0,31))/86400000;

 for(i=1900; i<2100 && offset>0; i++) { temp=lYearDays(i); offset-=temp; }

 if(offset<0) { offset+=temp; i--; }

 this.year = i;

 leap = leapMonth(i); //閏哪個月
 this.isLeap = false;

 for(i=1; i<13 && offset>0; i++) {
    //閏月
    if(leap>0 && i==(leap+1) && this.isLeap==false)
       { --i; this.isLeap = true; temp = leapDays(this.year); }
    else
       { temp = monthDays(this.year, i); }

    //解除閏月
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

//==============================傳回國曆 y年某m+1月的天數
function solarDays(y,m) {
 if(m==1)
    return(((y%4 == 0) && (y%100 != 0) || (y%400 == 0))? 29: 28);
 else
    return(solarMonth[m]);
}
//============================== 傳入 offset 傳回干支, 0=甲子
function cyclical(num) {
 return(Gan[num%10]+Zhi[num%12]);
}

//============================== 月曆屬性
function calElement(sYear,sMonth,sDay,week,lYear,lMonth,lDay,isLeap,cYear,cMonth,cDay) {

    this.isToday    = false;
    //國曆
    this.sYear      = sYear;   //西元年4位數字
    this.sMonth     = sMonth;  //西元月數字
    this.sDay       = sDay;    //西元日數字
    this.week       = week;    //星期, 1個中文
    //農曆
    this.lYear      = lYear;   //西元年4位數字
    this.lMonth     = lMonth;  //農曆月數字
    this.lDay       = lDay;    //農曆日數字
    this.isLeap     = isLeap;  //是否為農曆閏月?
    //八字
    this.cYear      = cYear;   //年柱, 2個中文
    this.cMonth     = cMonth;  //月柱, 2個中文
    this.cDay       = cDay;    //日柱, 2個中文

    this.color      = '';

    this.lunarFestival = ''; //農曆節日
    this.solarFestival = ''; //國曆節日
    this.solarTerms    = ''; //節氣
}

//===== 某年的第n個節氣為幾日(從0小寒起算)
function sTerm(y,n) {
 var offDate = new Date( ( 31556925974.7*(y-1900) + sTermInfo[n]*60000  ) + Date.UTC(1900,0,6,2,5) );
 return(offDate.getUTCDate());
}




//============================== 傳回月曆物件 (y年,m+1月)
/*
功能說明: 傳回整個月的日期資料物件

使用方式: OBJ = new calendar(年,零起算月);

OBJ.length      傳回當月最大日
OBJ.firstWeek   傳回當月一日星期

由 OBJ[日期].屬性名稱 即可取得各項值

OBJ[日期].isToday  傳回是否為今日 true 或 false

其他 OBJ[日期] 屬性參見 calElement() 中的註解
*/
function calendar(y,m) {

 var sDObj, lDObj, lY, lM, lD=1, lL, lX=0, tmp1, tmp2, tmp3;
 var cY, cM, cD; //年柱,月柱,日柱
 var lDPOS = new Array(3);
 var n = 0;
 var firstLM = 0;

 sDObj = new Date(y,m,1,0,0,0,0);    //當月一日日期

 this.length    = solarDays(y,m);    //國曆當月天數
 this.firstWeek = sDObj.getDay();    //國曆當月1日星期幾

 ////////年柱 1900年立春後為庚子年(60進制36)
 if(m<2) cY=cyclical(y-1900+36-1);
 else cY=cyclical(y-1900+36);
 var term2=sTerm(y,2); //立春日期

 ////////月柱 1900年1月小寒以前為 丙子月(60進制12)
 var firstNode = sTerm(y,m*2) //傳回當月「節」為幾日開始
 cM = cyclical((y-1900)*12+m+12);

 //當月一日與 1900/1/1 相差天數
 //1900/1/1與 1970/1/1 相差25567日, 1900/1/1 日柱為甲戌日(60進制10)
 var dayCyclical = Date.UTC(y,m,1,0,0,0,0)/86400000+25567+10;

 for(var i=0;i<this.length;i++) {

    if(lD>lX) {
       sDObj = new Date(y,m,i+1);    //當月一日日期
       lDObj = new Lunar(sDObj);     //農曆
       lY    = lDObj.year;           //農曆年
       lM    = lDObj.month;          //農曆月
       lD    = lDObj.day;            //農曆日
       lL    = lDObj.isLeap;         //農曆是否閏月
       lX    = lL? leapDays(lY): monthDays(lY,lM); //農曆當月最後一天

       if(n==0) firstLM = lM;
       lDPOS[n++] = i-lD+1;
    }

    //依節氣調整二月分的年柱, 以立春為界
    if(m==1 && (i+1)==term2) cY=cyclical(y-1900+36);
    //依節氣月柱, 以「節」為界
    if((i+1)==firstNode) cM = cyclical((y-1900)*12+m+13);
    //日柱
    cD = cyclical(dayCyclical+i);

    //sYear,sMonth,sDay,week,
    //lYear,lMonth,lDay,isLeap,
    //cYear,cMonth,cDay
    this[i] = new calElement(y, m+1, i+1, nStr1[(i+this.firstWeek)%7],
                             lY, lM, lD++, lL,
                             cY ,cM, cD );
 }

 //節氣
 tmp1=sTerm(y,m*2  )-1;
 tmp2=sTerm(y,m*2+1)-1;
 this[tmp1].solarTerms = solarTerm[m*2];
 this[tmp2].solarTerms = solarTerm[m*2+1];
 if(m==3) this[tmp1].color = 'red'; //清明顏色

 //國曆節日
 for(i in sFtv)
    if(sFtv[i].match(/^(\d{2})(\d{2})([\s\*])(.+)$/))
       if(Number(RegExp.$1)==(m+1)) {
          if(Number(RegExp.$2)<=this.length){
            this[Number(RegExp.$2)-1].solarFestival += RegExp.$4 + ' ';
            if(RegExp.$3=='*') this[Number(RegExp.$2)-1].color = 'red';
          }
       }

 //月週節日
 for(i in wFtv)
    if(wFtv[i].match(/^(\d{2})(\d)(\d)([\s\*])(.+)$/))
       if(Number(RegExp.$1)==(m+1)) {
          tmp1=Number(RegExp.$2);
          tmp2=Number(RegExp.$3);
          if(tmp1<5)
             this[((this.firstWeek>tmp2)?7:0) + 7*(tmp1-1) + tmp2 - this.firstWeek].solarFestival += RegExp.$5 + ' ';
          else {
             tmp1 -= 5;
             tmp3 = (this.firstWeek+this.length-1)%7; //當月最後一天星期?
             this[this.length - tmp3 - 7*tmp1 + tmp2 - (tmp2>tmp3?7:0) - 1 ].solarFestival += RegExp.$5 + ' ';
          }
       }

 //農曆節日
 for(i in lFtv)
    if(lFtv[i].match(/^(\d{2})(.{2})([\s\*])(.+)$/)) {
       tmp1=Number(RegExp.$1)-firstLM;
       if(tmp1==-11) tmp1=1;
       if(tmp1 >=0 && tmp1<n) {
          tmp2 = lDPOS[tmp1] + Number(RegExp.$2) -1;
          if( tmp2 >= 0 && tmp2<this.length && this[tmp2].isLeap!=true) {
             this[tmp2].lunarFestival += RegExp.$4 + ' ';
             if(RegExp.$3=='*') this[tmp2].color = 'red';
          }
       }
    }


 //復活節只出現在3或4月
 if(m==2 || m==3) {
    var estDay = new easter(y);
    if(m == estDay.m)
       this[estDay.d-1].solarFestival = this[estDay.d-1].solarFestival+' 復活節';
 }

 //黑色星期五
 if((this.firstWeek+12)%7==5)
    this[12].solarFestival += '黑色星期五';

 //今日
 if(y==tY && m==tM) this[tD-1].isToday = true;
}

//======================================= 傳回該年的復活節(春分後第一次滿月週後的第一主日)
function easter(y) {

 var term2=sTerm(y,5); //取得春分日期
 var dayTerm2 = new Date(Date.UTC(y,2,term2,0,0,0,0)); //取得春分的國曆日期物件(春分一定出現在3月)
 var lDayTerm2 = new Lunar(dayTerm2); //取得取得春分農曆

 if(lDayTerm2.day<15) //取得下個月圓的相差天數
    var lMlen= 15-lDayTerm2.day;
 else
    var lMlen= (lDayTerm2.isLeap? leapDays(y): monthDays(y,lDayTerm2.month)) - lDayTerm2.day + 15;

 //一天等於 1000*60*60*24 = 86400000 毫秒
 var l15 = new Date(dayTerm2.getTime() + 86400000*lMlen ); //求出第一次月圓為國曆幾日
 var dayEaster = new Date(l15.getTime() + 86400000*( 7-l15.getUTCDay() ) ); //求出下個週日

 this.m = dayEaster.getUTCMonth();
 this.d = dayEaster.getUTCDate();

}

//====================== 中文日期
function cDay(d){
 var s;

 switch (d) {
    case 10:
       s = '初十'; break;
    case 20:
       s = '二十'; break;
       break;
    case 30:
       s = '三十'; break;
       break;
    default :
       s = nStr2[Math.floor(d/10)];
       s += nStr1[d%10];
 }
 return(s);
}

///////////////////////////////////////////////////////////////////////////////

var cld;

function drawCld(SY,SM) {
 var i,sD,s,size;
 cld = new calendar(SY,SM);

 if(SY>1874 && SY<1909) yDisplay = '光緒' + (((SY-1874)==1)?'元':SY-1874);
 if(SY>1908 && SY<1912) yDisplay = '宣統' + (((SY-1908)==1)?'元':SY-1908);
 if(SY>1911) yDisplay = '民國' + (((SY-1911)==1)?'元':SY-1911);

 GZ.innerHTML = yDisplay +'年 農曆歲次' + cyclical(SY-1900+36) + '年 【'+Animals[(SY-4)%12]+'】';

 YMBG.innerHTML = "&nbsp;" + SY + "<BR>&nbsp;" + monthName[SM];

 for(i=0;i<42;i++) {

    sObj=eval('SD'+ i);
    lObj=eval('LD'+ i);

    sObj.className = '';

    sD = i - cld.firstWeek;

    if(sD>-1 && sD<cld.length) { //日期內
       sObj.innerHTML = sD+1;

       if(cld[sD].isToday) sObj.className = 'todyaColor'; //今日顏色

       sObj.style.color = cld[sD].color; //國定假日顏色

       if(cld[sD].lDay==1) //顯示農曆月
          if(cld[sD].isLeap) //閏月
            lObj.innerHTML = '<b>閏'+cld[sD].lMonth+'月' + (leapDays(cld[sD].lYear)==29?'小':'大')+'</b>';
          else //非閏月
            lObj.innerHTML = '<b>'+cld[sD].lMonth+'月' + (monthDays(cld[sD].lYear,cld[sD].lMonth)==29?'小':'大')+'</b>';
       else //顯示農曆日
          lObj.innerHTML = cDay(cld[sD].lDay);

       s=cld[sD].lunarFestival;
       if(s.length>0) { //農曆節日
          if(s.length>6) s = s.substr(0, 4)+'...';
          s = s.fontcolor('red');
       }
       else { //國曆節日
          s=cld[sD].solarFestival;
          if(s.length>0) {
             size = (s.charCodeAt(0)>0 && s.charCodeAt(0)<128)?8:4;
             if(s.length>size+2) s = s.substr(0, size)+'...';
             s=(s=='黑色星期五')?s.fontcolor('black'):s.fontcolor('blue');
          }
          else { //廿四節氣
             s=cld[sD].solarTerms;
             if(s.length>0) s = s.fontcolor('limegreen');
          }
       }

       if(cld[sD].solarTerms=='清明') s = '清明節'.fontcolor('red');



       if(s.length>0) lObj.innerHTML = s;

    }
    else { //非日期
       sObj.innerHTML = '';
       lObj.innerHTML = '';
    }
 }
}


function changeCld() {
 var y,m;
 y=CLD.SY.selectedIndex+1900;
 m=CLD.SM.selectedIndex;
 drawCld(y,m);
}

function pushBtm(K) {
 switch (K){
    case 'YU' :
       if(CLD.SY.selectedIndex>0) CLD.SY.selectedIndex--;
       break;
    case 'YD' :
       if(CLD.SY.selectedIndex<200) CLD.SY.selectedIndex++;
       break;
    case 'MU' :
       if(CLD.SM.selectedIndex>0) {
          CLD.SM.selectedIndex--;
       }
       else {
          CLD.SM.selectedIndex=11;
          if(CLD.SY.selectedIndex>0) CLD.SY.selectedIndex--;
       }
       break;
    case 'MD' :
       if(CLD.SM.selectedIndex<11) {
          CLD.SM.selectedIndex++;
       }
       else {
          CLD.SM.selectedIndex=0;
          if(CLD.SY.selectedIndex<200) CLD.SY.selectedIndex++;
       }
       break;
    default :
       CLD.SY.selectedIndex=tY-1900;
       CLD.SM.selectedIndex=tM;
 }
 changeCld();
}

var Today = new Date();
var tY = Today.getFullYear();
var tM = Today.getMonth();
var tD = Today.getDate();
//////////////////////////////////////////////////////////////////////////////

var width = "130";
var offsetx = 2;
var offsety = 8;

var x = 0;
var y = 0;
var snow = 0;
var sw = 0;
var cnt = 0;

var dStyle;
document.onmousemove = mEvn;

//顯示詳細日期資料
function mOvr(v) {
 var s,festival,spcday;
 var sObj=eval('SD'+ v);
 var d=sObj.innerHTML-1;

    //sYear,sMonth,sDay,week,
    //lYear,lMonth,lDay,isLeap,
    //cYear,cMonth,cDay

 if(sObj.innerHTML!='') {
    sObj.style.cursor = 's-resize';
    spcday=cld[d].sMonth==3 && cld[d].sDay==3*7?unescape('%u5F90%u7684%u751F%u65E5'):'';

    if(cld[d].solarTerms == '' && cld[d].solarFestival == '' && cld[d].lunarFestival == '' && spcday=='')
       festival = '';
    else
       festival = '<TABLE WIDTH=100% BORDER=0 CELLPADDING=2 CELLSPACING=0 BGCOLOR="#CCFFCC"><TR><TD>'+
       '<FONT COLOR="#000000" STYLE="font-size:9pt;">'+cld[d].solarTerms + ' ' + cld[d].solarFestival + ' ' + cld[d].lunarFestival+' '+spcday+'</FONT></TD>'+
       '</TR></TABLE>';

    s= '<TABLE WIDTH="130" BORDER=0 CELLPADDING="2" CELLSPACING=0 BGCOLOR="#000066" style="filter:Alpha(opacity=80)"><TR><TD>' +
       '<TABLE WIDTH=100% BORDER=0 CELLPADDING=0 CELLSPACING=0><TR><TD ALIGN="right"><FONT COLOR="#ffffff" STYLE="font-size:9pt;">'+
       cld[d].sYear+' 年 '+cld[d].sMonth+' 月 '+cld[d].sDay+' 日<br>星期'+cld[d].week+'<br>'+
       '<font color="violet">農曆'+(cld[d].isLeap?'閏 ':' ')+cld[d].lMonth+' 月 '+cld[d].lDay+' 日</font><br>'+
       '<font color="yellow">'+cld[d].cYear+'年 '+cld[d].cMonth+'月 '+cld[d].cDay + '日</font>'+
       '</FONT></TD></TR></TABLE>'+ festival +'</TD></TR></TABLE>';

    document.all["detail"].innerHTML = s;

    if (snow == 0) {
       dStyle.left = x+offsetx-(width/2);
       dStyle.top = y+offsety;
       dStyle.visibility = "visible";
       snow = 1;
    }
 }
}

//清除詳細日期資料
function mOut() {
 if ( cnt >= 1 ) { sw = 0; }
 if ( sw == 0 ) { snow = 0; dStyle.visibility = "hidden";}
 else cnt++;
}

//取得位置
function mEvn() {
 x=event.x;
 y=event.y;
 if (document.body.scrollLeft)
    {x=event.x+document.body.scrollLeft; y=event.y+document.body.scrollTop;}
 if (snow){
    dStyle.left = x+offsetx-(width/2);
    dStyle.top = y+offsety;
 }
}

/*****************************************************************************
                                世界時間計算
*****************************************************************************/

///////////////////////////////////////////////////////////////////////////
function initialize() {
 var key;

 //時間
 
 //月曆
 dStyle = detail.style;
 CLD.SY.selectedIndex=tY-1900;
 CLD.SM.selectedIndex=tM;
 drawCld(tY,tM);

}


function initializeTT() {
 var key;

 //時間
 map.filters.Light.Clear();
 map.filters.Light.addAmbient(255,255,255,60);
 map.filters.Light.addCone(120, 60, 80, 120, 60, 255,255,255,120,60);

 objContinentMenu=document.WorldClock.continentMenu;
 objCountryMenu=document.WorldClock.countryMenu;

 for (key in timeData)
    objContinentMenu[objContinentMenu.length]=new Option(key);


 var TZ1 = getCookie('TZ1');
 var TZ2 = getCookie('TZ2');


 if(TZ1=='') {TZ1=0; TZ2=18;}
 setTZ(TZ1,TZ2);

 tick();


 //月曆
 dStyle = detail.style;
 CLD.SY.selectedIndex=tY-1900;
 CLD.SM.selectedIndex=tM;
 drawCld(tY,tM);

}


//-->
</SCRIPT>

<STYLE>.todyaColor {
   BACKGROUND-COLOR: aqua
}
</STYLE>

<base target="_self">

<meta name="Microsoft Border" content="none, default">
</HEAD>
<BODY onload=initialize()  vlink="#0000FF">

<DIV id=detail
style="Z-INDEX: 3; FILTER: shadow(color=#333333,direction=135); WIDTH: 140px; POSITION: absolute; HEIGHT: 120px"></DIV>
<CENTER>
<div align="left">
<TABLE border=0>
<TBODY>
<TR><!-- //------------------------------ 世界時間 ----------------------------------->
    <!------------------------------ 萬年曆 ----------------------------------->
  <FORM name=CLD>
  <TD align=middle>
    <DIV style="Z-INDEX: -1; POSITION: absolute; TOP: 30px"><FONT id=YMBG
    style="FONT-SIZE: 100pt; COLOR: #f0f0f0; FONT-FAMILY: 'Arial Black'">&nbsp;0000<BR>&nbsp;JUN</FONT>
    </DIV>
    <TABLE border=0>
      <TBODY>
      <TR>
        <TD bgColor=#000080 colSpan=7><FONT style="FONT-SIZE: 9pt"
          color=#ffffff size=2>西曆<SELECT style="FONT-SIZE: 9pt"
          onchange=changeCld() name=SY>
            <SCRIPT language=JavaScript><!--
          for(i=1900;i<2101;i++) document.write('<option>'+i)
          //--></SCRIPT>
          </SELECT>年<SELECT style="FONT-SIZE: 9pt" onchange=changeCld()
          name=SM>
            <SCRIPT language=JavaScript><!--
          for(i=1;i<13;i++) document.write('<option>'+i)
          //--></SCRIPT>
          </SELECT>月</FONT> <FONT id=GZ face=標楷體 color=#ffffff
          size=4></FONT><BR></TD></TR>
      <TR align=middle bgColor=#e0e0e0>
        <TD width=54>日</TD>
        <TD width=54>一</TD>
        <TD width=54>二</TD>
        <TD width=50>三</TD>
        <TD width=54>四</TD>
        <TD width=54>五</TD>
        <TD width=54>六</TD></TR>
      <SCRIPT language=JavaScript><!--
          var gNum, color1, color2;

          // 星期六顏色
          switch (conWeekend) {
          case 1:
             color1 = 'black';
             color2 = color1;
             break;
          case 2:
             color1 = 'green';
             color2 = color1;
             break;
          case 3:
             color1 = 'red';
             color2 = color1;
             break;
          default :
             color1 = 'green';
             color2 = 'red';
          }

          for(i=0;i<6;i++) {
             document.write('<tr align=center>')
             for(j=0;j<7;j++) {
                gNum = i*7+j
               //document.write('<td id="GD' + gNum +'" onMouseOver="mOvr(' + gNum +')" onMouseOut="mOut()"><font id="SD' + gNum +'" size=5 face="Arial Black"')
                 document.write('<td id="GD' + gNum +'"><font id="SD' + gNum +'" size=5 face="Arial Black"')
              
                if(j == 0) document.write(' color=red')
                if(j == 6) {
                   if(i%2==1) document.write(' color='+color2)
                      else document.write(' color='+color1)
                }
                document.write(' TITLE=""> </font><br><font id="LD' + gNum + '" size=2 style="font-size:9pt"> </font></td>')
             }
             document.write('</tr>')
          }
          //--></SCRIPT>
      </TBODY></TABLE></TD>
  <TD vAlign=top align=middle width=40><BR><BR><BR>年<BR>
<BUTTON style="FONT-SIZE: 9pt" onclick="pushBtm('YD')">▲</BUTTON><BR>
<BUTTON style="FONT-SIZE: 9pt" onclick="pushBtm('YU')">▼</BUTTON>
<P>月<BR>
<BUTTON style="FONT-SIZE: 9pt" onclick="pushBtm('MD')">▲</BUTTON><BR>
<BUTTON style="FONT-SIZE: 9pt" onclick="pushBtm('MU')">▼</BUTTON><P>
<BUTTON style="FONT-SIZE: 9pt" onclick="pushBtm('')">當<BR>月</BUTTON><P>
</TD></FORM></TR></TBODY></TABLE></div>
</CENTER></BODY></HTML>
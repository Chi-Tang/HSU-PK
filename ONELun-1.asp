<!--
        ***************************************
         �A����&�@�ɮɶ� DHTML �{�� (�x�W��)
        ***************************************
                �̫�ק�: 2003/8/12


�p�G�zı�o�o�ӵ{�������A�z�i�H�ۥ���H���˪B�n�ͤ��ɡC�ۥѨ�
�νd��: �ǮաB�Ƿ|�B���|�B���q�����B�{����s�B�ӤH�����ѤH�d
�ߨϥΡC

Open Source ���N����ۧ@�v�A����Φ����ޥΩ�����e�ШӫH�i
���C�p�ݩ�u�ӷ~����Q�v�ت����ϥΦ��������{���X�θ�ơA�ݨ�
�o���H�ѭ����v�C

�̷s�����P��s��T�� http://sean.wox.org/ap/calendar/ ���G


�@                            �w��ӫH���۰Q�׬�s�P�����~��
                      �s���覡�Ghttp://sean.o4u.com/contact/
                      �@�@�@�@�@�@�@�@�@�@Sean Lin (�L����)
                          �L���L�H�Ч@�E�ФŧR�����ܧ󦹻���

-->

<HTML><HEAD><TITLE>�U�~��</TITLE>
<META content="�ɶ�;�@�ɮɶ�;�ɰ�;�A��;����;����;���;��k;�`��;������;�ɰ�;�`��;�G�Q�|�`��;�z��;�ͨv;world time clock;gregorian solar;chinese lunar;calendar" name=keywords>
<META content=All name=robots>
<META content="Gregorian Solar Calendar and Chinese Lunar Calendar" name=description>
<META http-equiv=Content-Type content="text/html; charset=big5">
<SCRIPT language=JScript>
<!--
/*****************************************************************************
                                 �ӤH���n�]�w
*****************************************************************************/

var conWeekend = 3;  // �P�����C�����: 1=�¦�, 2=���, 3=����, 4=�j�g��


/*****************************************************************************
                                 ������
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
var Gan=new Array("��","�A","��","�B","��","�v","��","��","��","��");
var Zhi=new Array("�l","��","�G","�f","��","�x","��","��","��","��","��","��");
var Animals=new Array("��","��","��","��","�s","�D","��","��","�U","��","��","��");
var solarTerm = new Array("�p�H","�j�H","�߬K","�B��","���h","�K��","�M��","�\�B","�߮L","�p��","�~��","�L��","�p��","�j��","�߬�","�B��","���S","���","�H�S","����","�ߥV","�p��","�j��","�V��");
var sTermInfo = new Array(0,21208,42467,63836,85337,107014,128867,150921,173149,195551,218072,240693,263343,285989,308563,331033,353350,375494,397447,419210,440795,462224,483532,504758);
var nStr1 = new Array('��','�@','�G','�T','�|','��','��','�C','�K','�E','�Q');
var nStr2 = new Array('��','�Q','��','��','�m');
var monthName = new Array("JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC");

//���`�� *��ܩ񰲤�
var sFtv = new Array(
"0101*���� ���إ���}�������",
"0111 �q�k�`",
"0115 �Įv�`",
"0123 �ۥѤ�",
"0204 �A���`",
"0214 ���H�`",
"0215 ���@�`",
"0219 �s�ͬ��B�ʬ�����",
"0228*�M��������",
"0229 ����",
"0301 �L�и`",
"0305 ���l�x�`",
"0308 ���k�`",
"0312 �Ӿ�` ����u�@������",
"0317 ����`",
"0320 �l�F�`",
"0321 ��H�`",
"0325 ���N�`",
"0326 �s���`",
"0329 �C�~�` ���R���P������",
"0330 �X���`",
"0401 �M�H�` �D�p�`",
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
"0606 �u�{�v�` ���Q�`",
"0609 �K���`",
"0615 ĵ��`",
"0630 �|�p�v�`",
"0701 �����` �����` �|�ȸ`",
"0706 �X�@�`",
"0711 ����`",
"0712 Ť�׸`",
"0808 ���˸`",
"0814 �ŭx�`",
"0827 �G���\\�Ϩ�",
"0901 �O�̸`",
"0903 �x�H�` �ܾԬ���",
"0909 ��|�` �߮v�`",
"0913 �k�ߤ�",
"0928 �Юv�` �դl�Ϩ�",
"1006 �ѤH�`",
"1010*��y������",
"1021 �ع��`",
"1025 �x�W���_�`",
"1031 �U�t�` �����Ϩ������� �a���`",
"1101 �ӤH�`",
"1111 �u�~�` �a�F�`",
"1117 �ۨӤ��`",
"1112 ����Ϩ������� ��v�` ���ؤ�ƴ_���`",
"1121 ���Ÿ`",
"1205 �����` ���H�`",
"1210 �H�v�`",
"1212 �˧L�`",
"1225 ��ˬ����� ���ڴ_���` �t�ϸ`",
"1227 �ؿv�v�`",
"1228 �q�H�`",
"1231 ���H�`");

//�Y�몺�ĴX�ӬP���X�C 5,6,7,8 ��ܨ�Ʋ� 1,2,3,4 �ӬP���X
var wFtv = new Array(
"0520 ���˸`",
"0716 �X�@�`",
"0730 �Q���а�a�g",
"1144 �P���`");

//�A��`��
var lFtv = new Array(
"0101*�K�`",
"0102*�^�Q�a",
"0103*����",
"0104 �ﯫ",
"0105 �}��",
"0109 �Ѥ���",
"0115 ���d�` �[���`",
"0202 �Y�� �g�a����",
"0323 ������",
"0408 �D���`",
"0505*�ݤȸ` �֤H�`",
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

//�@�ɮɶ����


/*****************************************************************************
                                    ����p��
*****************************************************************************/

//====================================== �Ǧ^�A�� y�~���`�Ѽ�
function lYearDays(y) {
 var i, sum = 348;
 for(i=0x8000; i>0x8; i>>=1) sum += (lunarInfo[y-1900] & i)? 1: 0;
 return(sum+leapDays(y));
}

//====================================== �Ǧ^�A�� y�~�|�몺�Ѽ�
function leapDays(y) {
 if(leapMonth(y)) return( (lunarInfo[y-1899]&0xf)==0xf? 30: 29);
 else return(0);
}

//====================================== �Ǧ^�A�� y�~�|���Ӥ� 1-12 , �S�|�Ǧ^ 0
function leapMonth(y) {
 var lm = lunarInfo[y-1900] & 0xf;
 return(lm==0xf?0:lm);
}

//====================================== �Ǧ^�A�� y�~m�몺�`�Ѽ�
function monthDays(y,m) {
 return( (lunarInfo[y-1900] & (0x10000>>m))? 30: 29 );
}


//====================================== ��X�A��, �ǤJ�������, �Ǧ^�A��������
//                                       �Ӫ����ݩʦ� .year .month .day .isLeap
function Lunar(objDate) {

 var i, leap=0, temp=0;
 var offset   = (Date.UTC(objDate.getFullYear(),objDate.getMonth(),objDate.getDate()) - Date.UTC(1900,0,31))/86400000;

 for(i=1900; i<2100 && offset>0; i++) { temp=lYearDays(i); offset-=temp; }

 if(offset<0) { offset+=temp; i--; }

 this.year = i;

 leap = leapMonth(i); //�|���Ӥ�
 this.isLeap = false;

 for(i=1; i<13 && offset>0; i++) {
    //�|��
    if(leap>0 && i==(leap+1) && this.isLeap==false)
       { --i; this.isLeap = true; temp = leapDays(this.year); }
    else
       { temp = monthDays(this.year, i); }

    //�Ѱ��|��
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

//==============================�Ǧ^��� y�~�Ym+1�몺�Ѽ�
function solarDays(y,m) {
 if(m==1)
    return(((y%4 == 0) && (y%100 != 0) || (y%400 == 0))? 29: 28);
 else
    return(solarMonth[m]);
}
//============================== �ǤJ offset �Ǧ^�z��, 0=�Ҥl
function cyclical(num) {
 return(Gan[num%10]+Zhi[num%12]);
}

//============================== ����ݩ�
function calElement(sYear,sMonth,sDay,week,lYear,lMonth,lDay,isLeap,cYear,cMonth,cDay) {

    this.isToday    = false;
    //���
    this.sYear      = sYear;   //�褸�~4��Ʀr
    this.sMonth     = sMonth;  //�褸��Ʀr
    this.sDay       = sDay;    //�褸��Ʀr
    this.week       = week;    //�P��, 1�Ӥ���
    //�A��
    this.lYear      = lYear;   //�褸�~4��Ʀr
    this.lMonth     = lMonth;  //�A���Ʀr
    this.lDay       = lDay;    //�A���Ʀr
    this.isLeap     = isLeap;  //�O�_���A��|��?
    //�K�r
    this.cYear      = cYear;   //�~�W, 2�Ӥ���
    this.cMonth     = cMonth;  //��W, 2�Ӥ���
    this.cDay       = cDay;    //��W, 2�Ӥ���

    this.color      = '';

    this.lunarFestival = ''; //�A��`��
    this.solarFestival = ''; //���`��
    this.solarTerms    = ''; //�`��
}

//===== �Y�~����n�Ӹ`�𬰴X��(�q0�p�H�_��)
function sTerm(y,n) {
 var offDate = new Date( ( 31556925974.7*(y-1900) + sTermInfo[n]*60000  ) + Date.UTC(1900,0,6,2,5) );
 return(offDate.getUTCDate());
}




//============================== �Ǧ^��䪫�� (y�~,m+1��)
/*
�\�໡��: �Ǧ^��Ӥ몺�����ƪ���

�ϥΤ覡: OBJ = new calendar(�~,�s�_���);

OBJ.length      �Ǧ^���̤j��
OBJ.firstWeek   �Ǧ^���@��P��

�� OBJ[���].�ݩʦW�� �Y�i���o�U����

OBJ[���].isToday  �Ǧ^�O�_������ true �� false

��L OBJ[���] �ݩʰѨ� calElement() ��������
*/
function calendar(y,m) {

 var sDObj, lDObj, lY, lM, lD=1, lL, lX=0, tmp1, tmp2, tmp3;
 var cY, cM, cD; //�~�W,��W,��W
 var lDPOS = new Array(3);
 var n = 0;
 var firstLM = 0;

 sDObj = new Date(y,m,1,0,0,0,0);    //���@����

 this.length    = solarDays(y,m);    //�����Ѽ�
 this.firstWeek = sDObj.getDay();    //�����1��P���X

 ////////�~�W 1900�~�߬K�ᬰ���l�~(60�i��36)
 if(m<2) cY=cyclical(y-1900+36-1);
 else cY=cyclical(y-1900+36);
 var term2=sTerm(y,2); //�߬K���

 ////////��W 1900�~1��p�H�H�e�� ���l��(60�i��12)
 var firstNode = sTerm(y,m*2) //�Ǧ^���u�`�v���X��}�l
 cM = cyclical((y-1900)*12+m+12);

 //���@��P 1900/1/1 �ۮt�Ѽ�
 //1900/1/1�P 1970/1/1 �ۮt25567��, 1900/1/1 ��W���Ҧ���(60�i��10)
 var dayCyclical = Date.UTC(y,m,1,0,0,0,0)/86400000+25567+10;

 for(var i=0;i<this.length;i++) {

    if(lD>lX) {
       sDObj = new Date(y,m,i+1);    //���@����
       lDObj = new Lunar(sDObj);     //�A��
       lY    = lDObj.year;           //�A��~
       lM    = lDObj.month;          //�A���
       lD    = lDObj.day;            //�A���
       lL    = lDObj.isLeap;         //�A��O�_�|��
       lX    = lL? leapDays(lY): monthDays(lY,lM); //�A����̫�@��

       if(n==0) firstLM = lM;
       lDPOS[n++] = i-lD+1;
    }

    //�̸`��վ�G������~�W, �H�߬K����
    if(m==1 && (i+1)==term2) cY=cyclical(y-1900+36);
    //�̸`���W, �H�u�`�v����
    if((i+1)==firstNode) cM = cyclical((y-1900)*12+m+13);
    //��W
    cD = cyclical(dayCyclical+i);

    //sYear,sMonth,sDay,week,
    //lYear,lMonth,lDay,isLeap,
    //cYear,cMonth,cDay
    this[i] = new calElement(y, m+1, i+1, nStr1[(i+this.firstWeek)%7],
                             lY, lM, lD++, lL,
                             cY ,cM, cD );
 }

 //�`��
 tmp1=sTerm(y,m*2  )-1;
 tmp2=sTerm(y,m*2+1)-1;
 this[tmp1].solarTerms = solarTerm[m*2];
 this[tmp2].solarTerms = solarTerm[m*2+1];
 if(m==3) this[tmp1].color = 'red'; //�M���C��

 //���`��
 for(i in sFtv)
    if(sFtv[i].match(/^(\d{2})(\d{2})([\s\*])(.+)$/))
       if(Number(RegExp.$1)==(m+1)) {
          if(Number(RegExp.$2)<=this.length){
            this[Number(RegExp.$2)-1].solarFestival += RegExp.$4 + ' ';
            if(RegExp.$3=='*') this[Number(RegExp.$2)-1].color = 'red';
          }
       }

 //��g�`��
 for(i in wFtv)
    if(wFtv[i].match(/^(\d{2})(\d)(\d)([\s\*])(.+)$/))
       if(Number(RegExp.$1)==(m+1)) {
          tmp1=Number(RegExp.$2);
          tmp2=Number(RegExp.$3);
          if(tmp1<5)
             this[((this.firstWeek>tmp2)?7:0) + 7*(tmp1-1) + tmp2 - this.firstWeek].solarFestival += RegExp.$5 + ' ';
          else {
             tmp1 -= 5;
             tmp3 = (this.firstWeek+this.length-1)%7; //���̫�@�ѬP��?
             this[this.length - tmp3 - 7*tmp1 + tmp2 - (tmp2>tmp3?7:0) - 1 ].solarFestival += RegExp.$5 + ' ';
          }
       }

 //�A��`��
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


 //�_���`�u�X�{�b3��4��
 if(m==2 || m==3) {
    var estDay = new easter(y);
    if(m == estDay.m)
       this[estDay.d-1].solarFestival = this[estDay.d-1].solarFestival+' �_���`';
 }

 //�¦�P����
 if((this.firstWeek+12)%7==5)
    this[12].solarFestival += '�¦�P����';

 //����
 if(y==tY && m==tM) this[tD-1].isToday = true;
}

//======================================= �Ǧ^�Ӧ~���_���`(�K����Ĥ@������g�᪺�Ĥ@�D��)
function easter(y) {

 var term2=sTerm(y,5); //���o�K�����
 var dayTerm2 = new Date(Date.UTC(y,2,term2,0,0,0,0)); //���o�K�������������(�K���@�w�X�{�b3��)
 var lDayTerm2 = new Lunar(dayTerm2); //���o���o�K���A��

 if(lDayTerm2.day<15) //���o�U�Ӥ�ꪺ�ۮt�Ѽ�
    var lMlen= 15-lDayTerm2.day;
 else
    var lMlen= (lDayTerm2.isLeap? leapDays(y): monthDays(y,lDayTerm2.month)) - lDayTerm2.day + 15;

 //�@�ѵ��� 1000*60*60*24 = 86400000 �@��
 var l15 = new Date(dayTerm2.getTime() + 86400000*lMlen ); //�D�X�Ĥ@����ꬰ���X��
 var dayEaster = new Date(l15.getTime() + 86400000*( 7-l15.getUTCDay() ) ); //�D�X�U�Ӷg��

 this.m = dayEaster.getUTCMonth();
 this.d = dayEaster.getUTCDate();

}

//====================== ������
function cDay(d){
 var s;

 switch (d) {
    case 10:
       s = '��Q'; break;
    case 20:
       s = '�G�Q'; break;
       break;
    case 30:
       s = '�T�Q'; break;
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

 if(SY>1874 && SY<1909) yDisplay = '����' + (((SY-1874)==1)?'��':SY-1874);
 if(SY>1908 && SY<1912) yDisplay = '�Ų�' + (((SY-1908)==1)?'��':SY-1908);
 if(SY>1911) yDisplay = '����' + (((SY-1911)==1)?'��':SY-1911);

 GZ.innerHTML = yDisplay +'�~ �A�䷳��' + cyclical(SY-1900+36) + '�~ �i'+Animals[(SY-4)%12]+'�j';

 YMBG.innerHTML = "&nbsp;" + SY + "<BR>&nbsp;" + monthName[SM];

 for(i=0;i<42;i++) {

    sObj=eval('SD'+ i);
    lObj=eval('LD'+ i);

    sObj.className = '';

    sD = i - cld.firstWeek;

    if(sD>-1 && sD<cld.length) { //�����
       sObj.innerHTML = sD+1;

       if(cld[sD].isToday) sObj.className = 'todyaColor'; //�����C��

       sObj.style.color = cld[sD].color; //��w�����C��

       if(cld[sD].lDay==1) //��ܹA���
          if(cld[sD].isLeap) //�|��
            lObj.innerHTML = '<b>�|'+cld[sD].lMonth+'��' + (leapDays(cld[sD].lYear)==29?'�p':'�j')+'</b>';
          else //�D�|��
            lObj.innerHTML = '<b>'+cld[sD].lMonth+'��' + (monthDays(cld[sD].lYear,cld[sD].lMonth)==29?'�p':'�j')+'</b>';
       else //��ܹA���
          lObj.innerHTML = cDay(cld[sD].lDay);

       s=cld[sD].lunarFestival;
       if(s.length>0) { //�A��`��
          if(s.length>6) s = s.substr(0, 4)+'...';
          s = s.fontcolor('red');
       }
       else { //���`��
          s=cld[sD].solarFestival;
          if(s.length>0) {
             size = (s.charCodeAt(0)>0 && s.charCodeAt(0)<128)?8:4;
             if(s.length>size+2) s = s.substr(0, size)+'...';
             s=(s=='�¦�P����')?s.fontcolor('black'):s.fontcolor('blue');
          }
          else { //�ܥ|�`��
             s=cld[sD].solarTerms;
             if(s.length>0) s = s.fontcolor('limegreen');
          }
       }

       if(cld[sD].solarTerms=='�M��') s = '�M���`'.fontcolor('red');



       if(s.length>0) lObj.innerHTML = s;

    }
    else { //�D���
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

//��ܸԲӤ�����
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
       cld[d].sYear+' �~ '+cld[d].sMonth+' �� '+cld[d].sDay+' ��<br>�P��'+cld[d].week+'<br>'+
       '<font color="violet">�A��'+(cld[d].isLeap?'�| ':' ')+cld[d].lMonth+' �� '+cld[d].lDay+' ��</font><br>'+
       '<font color="yellow">'+cld[d].cYear+'�~ '+cld[d].cMonth+'�� '+cld[d].cDay + '��</font>'+
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

//�M���ԲӤ�����
function mOut() {
 if ( cnt >= 1 ) { sw = 0; }
 if ( sw == 0 ) { snow = 0; dStyle.visibility = "hidden";}
 else cnt++;
}

//���o��m
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
                                �@�ɮɶ��p��
*****************************************************************************/

///////////////////////////////////////////////////////////////////////////
function initialize() {
 var key;

 //�ɶ�
 
 //���
 dStyle = detail.style;
 CLD.SY.selectedIndex=tY-1900;
 CLD.SM.selectedIndex=tM;
 drawCld(tY,tM);

}


function initializeTT() {
 var key;

 //�ɶ�
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


 //���
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
<TR><!-- //------------------------------ �@�ɮɶ� ----------------------------------->
    <!------------------------------ �U�~�� ----------------------------------->
  <FORM name=CLD>
  <TD align=middle>
    <DIV style="Z-INDEX: -1; POSITION: absolute; TOP: 30px"><FONT id=YMBG
    style="FONT-SIZE: 100pt; COLOR: #f0f0f0; FONT-FAMILY: 'Arial Black'">&nbsp;0000<BR>&nbsp;JUN</FONT>
    </DIV>
    <TABLE border=0>
      <TBODY>
      <TR>
        <TD bgColor=#000080 colSpan=7><FONT style="FONT-SIZE: 9pt"
          color=#ffffff size=2>���<SELECT style="FONT-SIZE: 9pt"
          onchange=changeCld() name=SY>
            <SCRIPT language=JavaScript><!--
          for(i=1900;i<2101;i++) document.write('<option>'+i)
          //--></SCRIPT>
          </SELECT>�~<SELECT style="FONT-SIZE: 9pt" onchange=changeCld()
          name=SM>
            <SCRIPT language=JavaScript><!--
          for(i=1;i<13;i++) document.write('<option>'+i)
          //--></SCRIPT>
          </SELECT>��</FONT> <FONT id=GZ face=�з��� color=#ffffff
          size=4></FONT><BR></TD></TR>
      <TR align=middle bgColor=#e0e0e0>
        <TD width=54>��</TD>
        <TD width=54>�@</TD>
        <TD width=54>�G</TD>
        <TD width=50>�T</TD>
        <TD width=54>�|</TD>
        <TD width=54>��</TD>
        <TD width=54>��</TD></TR>
      <SCRIPT language=JavaScript><!--
          var gNum, color1, color2;

          // �P�����C��
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
  <TD vAlign=top align=middle width=40><BR><BR><BR>�~<BR>
<BUTTON style="FONT-SIZE: 9pt" onclick="pushBtm('YD')">��</BUTTON><BR>
<BUTTON style="FONT-SIZE: 9pt" onclick="pushBtm('YU')">��</BUTTON>
<P>��<BR>
<BUTTON style="FONT-SIZE: 9pt" onclick="pushBtm('MD')">��</BUTTON><BR>
<BUTTON style="FONT-SIZE: 9pt" onclick="pushBtm('MU')">��</BUTTON><P>
<BUTTON style="FONT-SIZE: 9pt" onclick="pushBtm('')">��<BR>��</BUTTON><P>
</TD></FORM></TR></TBODY></TABLE></div>
</CENTER></BODY></HTML>
 
 <!-- #include virtual="/HSU-PK/DB.fun" -->
 
<%
mdbFile = "/HSU-PK/UsersPwd.mdb"
 '' mdbFile = "/Hmath/UsersPwd.mdb"
 ''mdbFile = "UsersPwd.mdb"
 mdbPassword = "kj6688"
 
 MySelf = Request.ServerVariables("PATH_INFO")
   Lesson = Request("Lesson")
   No = Request("No")
   No1 = Request("No1")
   No2 = Request("No2")

   'Name = Request("Name")
 Sex=Request("Sex")
SECTM=Request("SECTM")
TNUM=Request("TNUM")
DNUM=Request("DNUM")
HNUM=Request("HNUM")
LYR=Request("LYR")
 'SECTMS=SPLIT(SECTM, ",")
 'TNUMS=SPLIT(TNUM, ",") 
 'CNm=Request("NUM")
 ' 'RNDSN  SECTMS,TNUMS
   'For  k=0 to Ubound(SECTMS)
     ' RNDSN SECTMS(k),TNUMS(k)
   ' NEXT
 
ZK= Array("��","�A","��","�B","��","�v","��","��","��","��")
ZG= Array("�l","��","�G","�f","��","�w","��","��","��","��","��","��")

Mssg = "ok"
''On Error Resume Next 

If Request("Send") <> Empty Then
   SQL = "Select * From BIRTH " 
 Set rs = GetSecuredMdbRecordset( mdbFile, SQL, mdbPassword )
   ' SQL = "Select * From BIRTH " 
 'SQL = SQL & "Where �Ǹ�=" & No & " And �m�W='" & Name & "'"
 ' Set rs = GetMdbRecordset( "Testac-1.mdb", SQL )
  ' SQL = "Select * From ���Z�� " 
  ' SQL = SQL & "Where �Ǹ�=" & No & " And �m�W='" & Name & "'"
  ' Set rsScore = GetMdbRecordset( "Test.mdb", SQL )
  ''    SQLL = "Select * From "&Lesson&"k" 
  ''    Set rs = GetMdbRecordset( "Testac-1.mdb", SQLL )
     n=0
     TNum1=0
   
   ERNDSN  SECTM,TNUM,DNUM,HNUM
    ''RNDSN  SECTM,TNUM
    'For  k=0 to Ubound(SECTMS)
     ' RNDSN SECTMS(k),TNUMS(k)
    'NEXT
  YKKN=ERNDSN(SECTM, TNUM, DNUM, HNUM)
  YKKG=GRNDSN(SECTM, TNUM, DNUM, HNUM)
  LYGG=LRNDSN(LYR, "2", "20")

 ' Response.Write "<TR><TD>�~�z:"&YKKN & "</TD></TR>"
  
  '' On Error Resume Next 
  
  'Set conn = GetMdbConnection("Test1.mdb")
  Set cmd = Server.CreateObject( "ADODB.Command" )
  Set cmd.ActiveConnection = rs.ActiveConnection
  'SQLS ="Select * into ASP FROM ASPK"
  ' SQLS ="Select * into "&Lesson& " FROM " &Lesson&"K" �@
  ''SQLD ="Delete From "&Lesson
  ''cmd.CommandText = SQLD
    '   cmd.Execute
  '   
     ' Response.Redirect "TestKac-1t.asp?" & Request.QueryString
  Response.Redirect "Excer2-pg.asp?" & "YKK=" & YKKN & "&" & "YKG=" & YKKG& "&" & "LYG=" & LYGG & "&" & Request.QueryString
   
   '''Response.Redirect "Excel581-9.asp?"& "YKK="&YKKN& "&" & "TNUM=" & TNUM & "&" & "DNUM=" & DNUM & "&" & "HNUM=" & HNUM

End If    

 %> 

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<title></title>
</head>

<body background="../b01.jpg">
 
 
<%  '�p��~���z��
 FUNCTION ERNDSN(SECTM, TNUM, DNUM, HNUM)

' SUB ERNDSN(SECTM, TNUM, DNUM, HNUM)
  '  SQL = "Select * From BIRTH " 
 'SQL = SQL & "Where �Ǹ�=" & No & " And �m�W='" & Name & "'"
 ' Set rs = GetMdbRecordset( "Testac-1.mdb", SQL )
     ' TNUM1=30 
    ' TNUM1=TNUM1+TNUM 
     '  SECTM=SECTM+0
   D1=  #1912/2/18#
   D2=DateSerial(SECTM,TNUM,DNUM)
   DY=DateDiff("yyyy", D1, D2)
   DD=DateDiff("d", D1, D2) 
   YK8 = DY+8
   YK = YK8 MOD 10
   YG = DY MOD 12
   MG1= TNUM+1
   MG = MG1 MOD 12
   DK = DD MOD 10
   DG = DD MOD 12
  Response.Write "<TR><TD>���~:"&  D1 & "</TD></TR>"
  Response.Write "<TR><TD>�ͦ~:"&  D2 & "</TD></TR>"
 ' Response.Write "<TR><TD>�@�~:"&  DY & "</TD></TR>"
 ' Response.Write "<TR><TD>�@��:"&  DD & "</TD></TR>"
 ' Response.Write "<TR><TD>�~�z:"&YK & ZK(YK) & "</TD></TR>"
 ' Response.Write "<TR><TD>�~��:"&YG & ZG(YG) & "</TD></TR>"
 ' Response.Write "<TR><TD>��z:"&DK & ZK(DK) & "</TD></TR>"
 ' Response.Write "<TR><TD>���:"&DG & ZG(DG) & "</TD></TR>"
  Response.Write "<TR><TD>�ɤz:"&HNUM & "</TD></TR>"
    rs.AddNew
  'rs("�Ǹ�") = CLng(UserID)
  'rs("�Ǹ�") = CLng(Trim(Right(UserID,6)))
  'rs("�Ǹ�") = Trim(Right(UserID,6))
 rs("�Ǹ�") = Trim(Right(NO,12))
 rs("�m�W") =  Name
 rs("�~") = SECTM
 rs("��") = TNUM
 rs("��") = DNUM 
 rs("��") = HNUM
 rs("YK") = ZK(YK)
 rs("YG") = ZG(YG)
 rs("MG") = ZG(MG)
 rs("DK") = ZK(DK)
 rs("DG") = ZG(DG)
 rs("HG") = HNUM
 rs.Update

   ERNDSN=ZK(YK)
 END FUNCTION 
' END SUB
  %>    
<%  '�p��~��
 FUNCTION GRNDSN(SECTM, TNUM, DNUM, HNUM)

   D1=  #1912/2/18#
   D2=DateSerial(SECTM,TNUM,DNUM)
   DY=DateDiff("yyyy", D1, D2)
   DD=DateDiff("d", D1, D2) 
   YK8 = DY+8
   YK = YK8 MOD 10
   YG = DY MOD 12
   MG1= TNUM+1
   MG = MG1 MOD 12
   DK = DD MOD 10
   DG = DD MOD 12
  Response.Write "<TR><TD>���~:"&  D1 & "</TD></TR>"
  Response.Write "<TR><TD>�ͦ~:"&  D2 & "</TD></TR>"
 ' Response.Write "<TR><TD>�@�~:"&  DY & "</TD></TR>"
 ' Response.Write "<TR><TD>�@��:"&  DD & "</TD></TR>"
 ' Response.Write "<TR><TD>�~�z:"&YK & ZK(YK) & "</TD></TR>"
 Response.Write "<TR><TD>�~��:"&YG & ZG(YG) & "</TD></TR>"
 ' Response.Write "<TR><TD>��z:"&DK & ZK(DK) & "</TD></TR>"
 ' Response.Write "<TR><TD>���:"&DG & ZG(DG) & "</TD></TR>"
 ' Response.Write "<TR><TD>�ɤz:"&HNUM & "</TD></TR>"
   
 
 GRNDSN=ZG(YG)

 END FUNCTION 

  %>    
<%  '�p��y�~��
 FUNCTION LRNDSN(LYR, MON, DT)

   D1=  #1912/2/18#
   D2=DateSerial(LYR,MON, DT)
   DY=DateDiff("yyyy", D1, D2)
   DD=DateDiff("d", D1, D2) 
   YK8 = DY+8
   YK = YK8 MOD 10
   YG = DY MOD 12
   MG1= TNUM+1
   MG = MG1 MOD 12
   DK = DD MOD 10
   DG = DD MOD 12
  Response.Write "<TR><TD>���~:"&  D1 & "</TD></TR>"
  Response.Write "<TR><TD>�ͦ~:"&  D2 & "</TD></TR>"
 ' Response.Write "<TR><TD>�@�~:"&  DY & "</TD></TR>"
 ' Response.Write "<TR><TD>�@��:"&  DD & "</TD></TR>"
 ' Response.Write "<TR><TD>�~�z:"&YK & ZK(YK) & "</TD></TR>"
 Response.Write "<TR><TD>�~��:"&YG & ZG(YG) & "</TD></TR>"
 ' Response.Write "<TR><TD>��z:"&DK & ZK(DK) & "</TD></TR>"
 ' Response.Write "<TR><TD>���:"&DG & ZG(DG) & "</TD></TR>"
 ' Response.Write "<TR><TD>�ɤz:"&HNUM & "</TD></TR>"
   
 
 LRNDSN=ZG(YG)

 END FUNCTION 

  %>    
  <% '���հƵ{��2 
  FUNCTION HMCLD(rsTM) 
   Set fs = Server.CreateObject("Scripting.FileSystemObject")
     File = Server.MapPath(rsTM)
  Set txtf = fs.OpenTextFile( File )
If Not txtf.atEndOfStream Then	' ���T�w�٨S����F��������m
    Content = txtf.ReadAll	' Ū������ɮת����
   '' Lines = Replace(Content, vbCrLf, "<BR>" )
    ''Response.Write Lines
   ' Response.Write Content
End If

 HMCLD=Content
 
END FUNCTION
 

%>
 
   
   

<h2 align="center">�}�v���g�R���{��</h2>

<hr>

<h2>��ܩR�z���:</h2>

<blockquote>
    <form action="<%=Myself%>" method="GET"> 
        <p>��ءG<select name="Lesson" size="1"> 
            <option IsSelected("ZEWU", Lesson)>ZEWU</option> 
        </select></p> 
        <p>�Х���J�N���Ʀr�G<input type="text" size="20" name="No" Value="<%=No%>"></p>  
        <p>�п�J�Ĥ@(�U)�ơG<input type="integer" size="20" name="No1" Value="<%=No1%>"></p> 
        <p>�п�J�ĤG(�W)�ơG<input type="integer" size="20" name="No2" Value="<%=No2%>"></p> 
     <!---   �m�W�G<input type="text" size="20" name="Name"  Value="<%=Name%>"> --> 
    
        <p>�ʧO�G<select name="Sex" size="1">                                                                                                                                                                                                                                                                                                                 
            <option value="�k">�k</option><option value="�k">�k</option> 
             </select>
          </p>

           
      <h5> �A��褸?�~ �@ &���              &�鸹            &�ɨ�                      &���~�褸?�~ :</h5>                                                                                                                                                                                                               
         <p>�G<select name="SECTM" size="6"> 
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
          �G<select name="TNUM" size="6">                                                                                                                                                                                                                                                                                                                                                                  
            <option value="1">1</option><option value="2">2</option> 
             <option value="3">3</option><option value="4">4</option> 
             <option value="5">5</option><option value="6">6</option> 
             <option value="7">7</option><option value="8">8</option> 
             <option value="9">9</option><option value="10">10</option> 
             <option value="11">11</option><option value="12">12</option> 
            </select>
            <!.../p...>                                                                                                                                                                                                               
          �G<select name="DNUM" size="6">                                                                                                                                                                                                                                                                                                                                                                    
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
          �G<select name="HNUM" size="6">                                                                                                                                                                                                                                                                                                                                                                    
            <option value="�l">23~01</option><option value="��">01~03</option> 
             <option value="�G">03~05</option><option value="�f">05~07</option> 
             <option value="��">07~09</option><option value="�w">09~11</option> 
             <option value="��">11~13</option><option value="��">13~15</option> 
             <option value="��">15~17</option><option value="��">17~19</option> 
             <option value="��">19~21</option><option value="��">21~23</option> 
            </select>
                                                                                                                                                                               
         �G<select name="LYR" size="6">                           
             <option value="2007">2007</option><option value="2008">2008</option>  <option value="2009">2009</option><option value="2010">2010</option>
              <option value="2011">2011</option><option value="2012">2012</option><option value="2013">2013</option><option value="2014">2014</option>
             <option value="2015">2015</option><option value="2016">2016</option><option value="20176">2017</option><option value="2018">2018</option>
             <option value="2019">2019</option>  <option value="2020">2020</option><option value="2021">2021</option>
                <!.../p...>     
            </select>                            
          </p> 
         <p><input type="submit" Name="Send" value="�w��i�J��a"> </p> 
         <p>�@ </p> 
    </form> 
  </blockquote>  
 
<hr> 
<FONT Color=Red><%=Msg%></FONT> 
</boody> 
</html>

<%  
Function IsSelected( Which, Lesson ) 
   If Which = Lesson Then IsSelected = Selected 
End Function 
%>
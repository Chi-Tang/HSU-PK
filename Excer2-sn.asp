<!-- #include virtual="/HSU-PK/DB.fun" -->
<%
  Sex = Request("Sex")
  SECTM=Request("SECTM")
TNUM=Request("TNUM")
DNUM=Request("DNUM")
HNUM=Request("HNUM")
LYR=Request("LYR")
 YKK=Request("YKK")
 YKG=Request("YKG")
 LYG=Request("LYG")
ReDim A(220)
ReDim B(220)
ReDim AW(220)
ReDim AKI(220)

  'SQL = "Select * From [B2:N12]"
 SQL = "Select * From 命宮表"
Set rs3 = GetExcelRecordset( "Excel08.xls",  SQL)
If rs3 Is Nothing Then
    Response.Write "GetExcelRecordset 呼叫失敗!"
    Response.End
End If 
  ReDim z3(16)
  z3(0)="命宮"
  'z3k=" "
   wk=" "
   ZLG= Array("子","丑","寅","卯","辰","已","午","未","申","酉","戌","亥")
   ZLN= Array("貪狼","巨門","祿存","文曲","廉貞","武曲","破軍","武曲","廉貞","文曲","祿存","巨門")
   ZBN= Array("火星","天相","天梁","天同","文昌","天機","火星","天相","天梁","天同","文昌","天機")

 %>
<HTML>
<BODY bgcolor="#FFFFFF">
  <CENTER><H2>紫 微 斗 數 命 盤 與 評 論 表 如 下<HR></H2>
<TABLE BORDER=1>
<TR BGCOLOR=#00FFFF>
 </TR>
 
<%
  ' Part II：輸出資料表的「內容」
  rs3.MoveFirst		' 將目前資料錄移到第一筆
  While Not rs3.EOF	' 判斷是否過了最後一筆
    IF Trim(rs3(0))= Trim(TNUM)  Then
   'IF rs3(0)= "9"  Then 
      'ReDim z3(16)
          ' z3(0)=rs3(0)    
       For i=1 to rs3.Fields.Count-1
        if Trim(rs3(i).Name)= Trim(HNUM) then
        'if rs3(i).Name= "辰" then
           z3(i)=rs3(i)
           z3k =rs3(i)
          ' z3(0)="(身)命宮"+z3k
          ' z30k="(身)命宮"
          ' z3(i) =  rs3(i)+rs3(i).Name
         ''rs3(i) =rs3(i).Name+rs3(i)
        ''' Exit For
        End If 
       Next
    ' rs3.Update
    End If  
     rs3.MoveNext	' 移到下一筆
  Wend
 
 A(5)=z3(0)
 B(5)=z3k
 AW(5)=z3(0)&wk
 AKI(5)="K5"
 
  For n=0 to 11
        if z3k=ZLG(n) then
             MYL=ZLN(n)
          end if 
      Next   
  
   For n=0 to 11
        if Trim(YKG)=ZLG(n) then
           ''zgtk1=n
          MYB=ZBN(n)
         end if 
      Next   
 
 
 %>

<%
 'SQL = "Select * From [B2:N12]"
 SQL = "Select * From 身宮表"
Set rsb1 = GetExcelRecordset( "Excel08.xls",  SQL)
If rsb1 Is Nothing Then
    Response.Write "GetExcelRecordset 呼叫失敗!"
    Response.End
End If 
  ReDim zb1(16)
  zb1(0)="身宮"
  'zb1k=" "
   wk=" "
   
  ' Part II：輸出資料表的「內容」
  rsb1.MoveFirst		' 將目前資料錄移到第一筆
  While Not rsb1.EOF	' 判斷是否過了最後一筆
    IF Trim(rsb1(0))= Trim(TNUM)  Then
   'IF rsb1(0)= "9"  Then 
       For i=1 to rsb1.Fields.Count-1
        if Trim(rsb1(i).Name)= Trim(HNUM) then
        'if rsb1(i).Name= "辰" then
           zb1(i)=rsb1(i)
           zb1k =rsb1(i)
         End If 
       Next
     'rsb1.Update
    End If  
     rsb1.MoveNext	' 移到下一筆
  Wend
 
 A(3)=zb1(0)
 B(3)=zb1k
 AW(3)=zb1(0)&wk
 AKI(3)="K3"
 
 %>
<%'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From 餘宮表"
Set rs5 = GetExcelRecordset( "Excel08.xls",  SQL)
If rs5 Is Nothing Then
    Response.Write "GetExcelRecordset 呼叫失敗!"
    Response.End
End If 
  ReDim z5(16)
 ' z5(0)="身命宮"
  z5k=" "
  k=5
  wk=" "

%>
<%
  '' Part I：輸出「抬頭名稱」
  'For i=0 to rs5.Fields.Count-1
  '  Response.Write "<TD>" & rs5(i).Name & "</TD>"
 ' Next
%>

<%
  ' Part II：輸出資料表的「內容」
  rs5.MoveFirst		' 將目前資料錄移到第一筆
  While Not rs5.EOF	' 判斷是否過了最後一筆
         k=k+1
   ' For j=1 to 11 
    ' IF rs5(0)= "十"  Then 
      'ReDim z5(16)
      z5(0)=rs5(0)
       For i=1 to rs5.Fields.Count-1
         if rs5(i).Name= z3k then
          'if rs5(i).Name= "寅" then
           'z5(0)="身命宮"
           z5(i)=rs5(i)
           z5k =rs5(i)
          ' z5(i) =  rs5(i)+rs5(i).Name
         ''rs5(i) =rs5(i).Name+rs5(i)
        ''' Exit For
        End If 
       Next
    ' rs5.Update
   ' End If  
   ' rs5.MoveNext	' 移到下一筆
 ' Wend  
  A(0+k)=z5(0)
  B(0+k)=z5k
  AW(0+k)=z5(0)&wk
  AKI(0+k)="K"&k
 
    
  rs5.MoveNext	' 移到下一筆
 Wend

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
  'SQL = "Select * From [B2:N12]"
 SQL = "Select * From 五行局"
Set rs2 = GetExcelRecordset( "Excel08.xls",  SQL)
If rs2 Is Nothing Then
    Response.Write "GetExcelRecordset 呼叫失敗!"
    Response.End
End If 
  ReDim z2(16)
  z2(0)="(身)命宮"
 '' z2k=" "
  wk=" "

%>


<%
  ' Part II：輸出資料表的「內容」
  rs2.MoveFirst		' 將目前資料錄移到第一筆
  While Not rs2.EOF	' 判斷是否過了最後一筆
    IF Trim(rs2(0))= Trim(YKK)  Then 
   
    '  IF rs2(0)= "戊"  Then 
      'ReDim z2(16)
         ' z2(0)=rs2(0)
       For i=1 to rs2.Fields.Count-1
          if rs2(i).Name= z3k then
          ' if rs2(i).Name= "寅" then
           'z2(0)="(身)命宮"
           z2(0)=rs2(i)
          ' z2(i)=rs2(i)
           z2k =rs2(i)
          ' z2(i) =  rs2(i)+rs2(i).Name
         ''rs2(i) =rs2(i).Name+rs2(i)
        ''' Exit For
        End If 
       Next
     rs2.Update
    End If  
     rs2.MoveNext	' 移到下一筆
  Wend
  
 'For j=o to UBound(z2)
   ' WREXCE z2(0), z2(j), z2K ; SUB WREXCE(zn0, zni, znk) ;END SUB 
  ' Next 
  A(2)=z2(0)
  B(2)=z2k
  AW(2)=z2(0)&wk
  AKI(2)= "K2"

 %>  


<%
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'SQL = "Select * From [B2:N12]"
 SQL = "Select * From 紫微星"
Set rs1 = GetExcelRecordset( "Excel08.xls",  SQL)
If rs1 Is Nothing Then
    Response.Write "GetExcelRecordset 呼叫失敗!"
    Response.End
End If 
  ReDim z1(16)
  z1(0)="紫微星"
  z1k=" "
  z11k=" "
  ZB= Array("子","丑","寅","卯","辰","已","午","未","申","酉","戌","亥")
  
  ZF= Array("辰","卯","寅","丑","子","亥","戌","酉","申","未","午","已")

%>

<%
  ' Part II：輸出資料表的「內容」
  rs1.MoveFirst		' 將目前資料錄移到第一筆
  While Not rs1.EOF	' 判斷是否過了最後一筆
    IF Trim(rs1(0))= Trim(DNUM)  Then 
    'IF rs1(0)= "8"  Then 
      'ReDim z1(16)
     
       For i=1 to rs1.Fields.Count-1
        if rs1(i).Name= z2k then
         'if rs1(i).Name= "木三局" then
           '''z1(0)="(紫微)天府星"
           z1(i)=rs1(i)
           z1k =rs1(i)
          For j=0 to 11
            if ZB(j)= rs1(i) then
               z11k = ZF(j)
            End If 
         Next

         ' z1(i) =  rs1(i)+rs1(i).Name
         ''rs1(i) =rs1(i).Name+rs1(i)
        ''' Exit For
        End If 
       Next
     rs1.Update
    End If  
     rs1.MoveNext	' 移到下一筆
  Wend 
%> 

<%
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From 甲星盤"
Set rs = GetExcelRecordset( "Excel08.xls",  SQL)
If rs Is Nothing Then
    Response.Write "GetExcelRecordset 呼叫失敗!"
    Response.End
End If 
  ReDim z(16)
 ' z(0)="(紫微)天府宮"
  zk=" "
  k=20
%>


<%
  ' Part II：輸出資料表的「內容」
  rs.MoveFirst		' 將目前資料錄移到第一筆
  While Not rs.EOF	' 判斷是否過了最後一筆
      k=k+1
   ' For j=1 to 11 
   ' IF rs(0)= "十"  Then 
          z(0)=rs(0)
         ' zzo=rs(0)
       For i=1 to rs.Fields.Count-1
        if rs(i).Name= z1k then
         'if rs(i).Name= "申" then
           'z(0)="(紫微)天府宮"
           z(i)=rs(i)
           zk =rs(i)
         End If 
       Next
    ' rs.Update
   ' End If  
   ' rs.MoveNext	' 移到下一筆
 ' Wend 
 
 
  ''SQL = "Select * From 甲星旺表"([B2:N12]") 
   ' WKW=WRNDSN(zzo, zk)
   wkw=WRNDSN(z(0), zk)
  
  A(0+k)=z(0)
  B(0+k)=zk
  AW(0+k)=z(0)&wkw
    Select Case Trim(wkw)
     Case "廟","旺"
        AKI(0+k)="S"&k&"A"
     Case "得","陷"
        AKI(0+k)="S"&k&"B"
    Case Else
        AKI(0+k)="S"&k&"C"
    End  Select      
 
 rs.MoveNext	' 移到下一筆
 Wend
  
%>

<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
FUNCTION WRNDSN(zzo, zk)
'' WRNDSN=WK  END FUNCTION 
 SQL = "Select * From 甲星旺表"
Set rsd = GetExcelRecordset( "Excel08.xls",  SQL)
 'If rsd Is Nothing Then
 '   Response.Write "GetExcelRecordset 呼叫失敗!"
  '  Response.End
 'End If 
  ReDim w(16)
  w(0)="旺度"  
  wk=" "

  ' Part II：輸出資料表的「內容」
  rsd.MoveFirst		' 將目前資料錄移到第一筆
  While Not rsd.EOF	' 判斷是否過了最後一筆
    IF rsd(0)= zzo  Then
    'IF rsd(0)= z(0)  Then
    'IF rsd(0)= "(紫微)天府"  Then 
      'ReDim z(16)
     
       For i=1 to rsd.Fields.Count-1
        if rsd(i).Name=  zk then
          'if rsd(i).Name= "寅" then
          ' w(0)="旺度"
           w(i)=rsd(i)
           wk =rsd(i)
         End If 
       Next
     rsd.Update
    End If  
     rsd.MoveNext	' 移到下一筆
  Wend 
 WRNDSN=wk  
END FUNCTION 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

SUB  WRNDSNft(zzo, zftk)

  
  ''SQL = "Select * From [B2:N12]"
 SQL = "Select * From 甲星旺表"
Set rsdt = GetExcelRecordset( "Excel09.xls",  SQL)
 'If rsdt Is Nothing Then
 '   Response.Write "GetExcelRecordset 呼叫失敗!"
  '  Response.End
 'End If 
  ReDim w(16)
  w(0)="旺度"
  wk=" "

  ' Part II：輸出資料表的「內容」
  rsdt.MoveFirst		' 將目前資料錄移到第一筆
  While Not rsdt.EOF	' 判斷是否過了最後一筆
    IF rsd(0)= zzo  Then
    'IF rsdt(0)= zft(0)  Then
    'IF rsdt(0)= "(紫微)天府"  Then 
      'ReDim zft(16)
     
       For i=1 to rsdt.Fields.Count-1
        if rsdt(i).Name=  zftk then
          'if rsdt(i).Name= "寅" then
          ' w(0)="旺度"
           w(i)=rsdt(i)
           wk =rsdt(i)
         End If 
       Next
     rsdt.Update
    End If  
     rsdt.MoveNext	' 移到下一筆
  Wend 
   A(0+k)=zft(0)
  B(0+k)=zftk
  AW(0+k)=zft(0)& wk
  'AW(0+k)=zft(0)&wk
 
  AW(0+k)=zft(0)&wk
    Select Case Trim(wk)
     Case "廟","旺"
        AKI(0+k)="S"&k&"A"
     Case "得","陷"
        AKI(0+k)="S"&k&"B"
    Case Else
        AKI(0+k)="S"&k&"C"
    End  Select      
 
  
  END SUB
%>

<%
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
 ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From 甲星盤2"
Set rsf = GetExcelRecordset( "Excel08.xls",  SQL)
If rsf Is Nothing Then
    Response.Write "GetExcelRecordset 呼叫失敗!"
    Response.End
End If 
  ReDim zf(16)
 ' zf(0)="(紫微)天府宮"
  zfk=" "
  k=30
%>


<%
  ' Part II：輸出資料表的「內容」
  rsf.MoveFirst		' 將目前資料錄移到第一筆
  While Not rsf.EOF	' 判斷是否過了最後一筆
     k=k+1
   ' For j=1 to 11 
   ' IF rsf(0)= "十"  Then 
      'ReDim zf(16)
     zf(0)=rsf(0)
       For i=1 to rsf.Fields.Count-1
        if rsf(i).Name= z11k then
         'if rsf(i).Name= "申" then
           'zf(0)="(紫微)天府宮"
           zf(i)=rsf(i)
           zfk =rsf(i)
         End If 
       Next
    ' rsf.Update
   ' End If  
   ' rsf.MoveNext	' 移到下一筆
 ' Wend 
  ''SQL = "Select * From 甲星旺表"([B2:N12]") 
   ' WKW=WRNDSN(zzo, zfk)
   wkw=WRNDSN(zf(0), zfk)
  
  A(0+k)=zf(0)
  B(0+k)=zfk
  AW(0+k)=zf(0)&wkw
 
  Select Case Trim(wkw)
     Case "廟","旺"
        AKI(0+k)="T"&k&"A"
     Case "得","陷"
        AKI(0+k)="T"&k&"B"
    Case Else
        AKI(0+k)="T"&k7&"C"
    End  Select      
 
  
 rsf.MoveNext	' 移到下一筆
 Wend
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Select Case Trim(YKK)
     Case "甲","丙","戊","庚","壬"
         SexYk="陽"&Sex
           Select Case Trim(Sex)
             Case "男"
                 SexYKT="S"
             Case "女"
                 SexYKT="R"
            ' Case Else
              ' SexYKT="N""
             End  Select   
         
     Case "乙","丁","己","辛","癸"
         SexYk="陰"&Sex
            Select Case Trim(Sex)
             Case "男"
                 SexYKT="R"
             Case "女"
                 SexYKT="S"
            ' Case Else
              ' SexYKT="N""
             End  Select 
     Case Else
         SexYk="中"&Sex
    End  Select   
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From 甲星盤3"
Set rsft = GetExcelRecordset( "Excel09.xls",  SQL)
If rsft Is Nothing Then
    Response.Write "GetExcelRecordset 呼叫失敗!"
    Response.End
End If 
  ReDim zft(16)
 ' zft(0)="祿存星"
  zftk=" "
  k=40


  ' Part II：輸出資料表的「內容」
  rsft.MoveFirst		' 將目前資料錄移到第一筆
  While Not rsft.EOF	' 判斷是否過了最後一筆
     k=k+1
   ' For j=1 to 11 
   ' IF rsft(0)= "十"  Then 
        zft(0)=rsft(0)
       For i=1 to rsft.Fields.Count-1
        if rsft(i).Name= Trim(YKK) then
         'if rsft(i).Name= "申" then
           'zft(0)="祿存星"
           zft(i)=rsft(i)
           zftk =rsft(i)
           
          IF Trim(rsft(0))= "祿存"  Then 
             subtk=zftk 
              
             ' DOCTOR SexYKT, subtk
           End If
       End If  
     Next
          
    ' rsft.MoveNext	' 移到下一筆
  ' Wend 
  ''SQL = "Select * From 甲星旺表"([B2:N12]") 
   ' WKW=WRNDSN(zzo, zftk)
   wkw=WRNDSN(zft(0), zftk)
    
  A(0+k)=zft(0)
  B(0+k)=zftk
  AW(0+k)=zft(0)& wkw
  'AW(0+k)=zft(0)&wkw
   
 rsft.MoveNext	' 移到下一筆
 Wend
 
  DOCTOR SexYKT, subtk

   %>
 
 <%
 
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 SUB DOCTOR(SexYKT, zftk)
 
 ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From 博士表"
Set rspt = GetExcelRecordset( "Excel09.xls",  SQL)
If rspt Is Nothing Then
    Response.Write "GetExcelRecordset 呼叫失敗!"
    Response.End
End If 
  ReDim zpt(24)
 ' zpt(0)="祿存星"
   wk=" "
  k=75


  ' Part II：輸出資料表的「內容」
  rspt.MoveFirst		' 將目前資料錄移到第一筆
  While Not rspt.EOF	' 判斷是否過了最後一筆
     k=k+1
   ' For j=1 to 11 
   ' IF rspt(0)= "十"  Then 
      'ReDim zpt(24)
     zpt(0)=rspt(0)
       For i=1 to rspt.Fields.Count-1
        if rspt(i).Name= Trim(SexYKT&zftk) then
         'if rspt(i).Name= "申" then
           'zpt(0)="祿存星"
           zpt(i)=rspt(i)
           zptk =rspt(i)
         End If 
       
       Next
          
    ' rspt.MoveNext	' 移到下一筆
  ' Wend  
 
  A(0+k)=zpt(0)
  B(0+k)=zptk
  AW(0+k)=zpt(0)&wk
  'AW(0+k)=zpt(0)&wk

    rspt.MoveNext	' 移到下一筆
   Wend
 
 END SUB
 
  %>
 
 
 <% 
 'SQL = "Select * From [B2:N12]"
 SQL = "Select * From 火星表"
Set rs6 = GetExcelRecordset( "Excel09.xls",  SQL)
If rs6 Is Nothing Then
    Response.Write "GetExcelRecordset 呼叫失敗!"
    Response.End
End If 
  ReDim z6(16)
  z6(0)="火星"
  z6k=" "
  k=51

  ' Part II：輸出資料表的「內容」
  rs6.MoveFirst		' 將目前資料錄移到第一筆
  While Not rs6.EOF	' 判斷是否過了最後一筆
       '' z6(0)=rs6(0)
       z6(0)="火星"
   IF Trim(rs6(0))= Trim(YKG)  Then 
    'IF rs6(0)= "8"  Then 
          
       For i=1 to rs6.Fields.Count-1
        if rs6(i).Name= Trim(HNUM) then
         'if rs6(i).Name= "木三局" then
           
           z6(i)=rs6(i)
           z6k =rs6(i)
         
        End If 
       Next
     rs6.Update
    End If  
     rs6.MoveNext	' 移到下一筆
  Wend 
%> 

<%
    
  ''SQL = "Select * From 甲星旺表"([B2:N12]") 
   ' WKW=WRNDSN(zzo, zftk)
   wkw=WRNDSN(z6(0), z6k)
    
  A(0+k)=z6(0)
  B(0+k)=z6k
  AW(0+k)=z6(0)&wkw
      
%>


<% 
 'SQL = "Select * From [B2:N12]"
 SQL = "Select * From 鈴星表"
Set rs7 = GetExcelRecordset( "Excel09.xls",  SQL)
If rs7 Is Nothing Then
    Response.Write "GetExcelRecordset 呼叫失敗!"
    Response.End
End If 
  ReDim z7(16)
  z7(0)="鈴星"
  z7k=" "
  k=52

  ' Part II：輸出資料表的「內容」
  rs7.MoveFirst		' 將目前資料錄移到第一筆
  While Not rs7.EOF	' 判斷是否過了最後一筆
       '' z7(0)=rs7(0)
       z7(0)="鈴星"
   IF Trim(rs7(0))= Trim(YKG)  Then 
    'IF rs7(0)= "8"  Then 
          
       For i=1 to rs7.Fields.Count-1
        if rs7(i).Name= Trim(HNUM) then
         'if rs7(i).Name= "木三局" then
           
           z7(i)=rs7(i)
           z7k =rs7(i)
         
        End If 
       Next
     rs7.Update
    End If  
     rs7.MoveNext	' 移到下一筆
  Wend 
%> 

<%
    
  ''SQL = "Select * From 甲星旺表"([B2:N12]") 
   ' WKW=WRNDSN(zzo, zftk)
   wkw=WRNDSN(z7(0), z7k)
    
  A(0+k)=z7(0)
  B(0+k)=z7k
  AW(0+k)=z7(0)&wkw
      
%>


<%
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From 甲星盤4"
Set rsf7 = GetExcelRecordset( "Excel09.xls",  SQL)
If rsf7 Is Nothing Then
    Response.Write "GetExcelRecordset 呼叫失敗!"
    Response.End
End If 
  ReDim zf7(16)
 ' zf7(0)="文昌(曲)星"
  zf7k=" "
  k=53

  ' Part II：輸出資料表的「內容」
  rsf7.MoveFirst		' 將目前資料錄移到第一筆
  While Not rsf7.EOF	' 判斷是否過了最後一筆
     k=k+1
     zf7(0)=rsf7(0)
       For i=1 to rsf7.Fields.Count-1
        if rsf7(i).Name= Trim(HNUM) then
         'if rsf7(i).Name= "申" then
           'zf7(0)="(紫微)天府宮"
           zf7(i)=rsf7(i)
           zf7k =rsf7(i)
         End If 
       Next
    ' rsf7.Update
   ' End If  
   ' rsf7.MoveNext	' 移到下一筆
 ' Wend  
   ''SQL = "Select * From 甲星旺表"([B2:N12]") 
   ' WKW=WRNDSN(zzo, zf7k)
   wkw=WRNDSN(zf7(0), zf7k)
    
  A(0+k)=zf7(0)
  B(0+k)=zf7k
  AW(0+k)=zf7(0)&wkw
 
 rsf7.MoveNext	' 移到下一筆
 Wend
 
  %>
  
  <%

 
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From 年支星表"
Set rsgt = GetExcelRecordset( "Excel09.xls",  SQL)
If rsgt Is Nothing Then
    Response.Write "GetExcelRecordset 呼叫失敗!"
    Response.End
End If 
  ReDim zgt(16)
 ' zgt(0)="天才壽"
  zgtk=" "
  wk=" "
  k=90


  ' Part II：輸出資料表的「內容」
  rsgt.MoveFirst		' 將目前資料錄移到第一筆
  While Not rsgt.EOF	' 判斷是否過了最後一筆
     k=k+1
     zgt(0)=rsgt(0)
       For i=1 to rsgt.Fields.Count-1
       if rsgt(i).Name= Trim(YKG) then
         'if rsgt(i).Name= "申" then
           'zgt(0)="天才壽"
           zgt(i)=rsgt(i)
           zgtk =rsgt(i)
     '   Select Case Trim(rsgt(0))
    ' Case "天才"
     '      For n=5 to 17
      '         if rsgt(i)=A(n) then
      '            zgtk=B(n)
       '        end if                
        '     Next

     ' Case "天壽" 
     '      BN5=Trim(B(3))
     '      YKGS=Trim(YKG)
      '   'zgtk= GGRNDSN(B(3), YKG)
      '   GGRNDSN BN5, YKGS

    '  End  Select      
        
         IF Trim(rsgt(0))= "天才"  Then 
            For n=5 to 17
               if rsgt(i)=A(n) then
                  zgtk=B(n)
               end if                
             Next
          End If
           IF Trim(rsgt(0))= "天壽"  Then 
               zgtk= GGRNDSN(B(3), YKG)
                BN5=Trim(B(3))
                YKGS=Trim(YKG)
              ' zgtk= GGRNDSN(BN3, YKGS)
             ' GGRNDSN BN3, YKGS

            End IF
       End If  
     Next
          
    ' rsgt.MoveNext	' 移到下一筆
  ' Wend 
 
  A(0+k)=zgt(0)
  B(0+k)=zgtk
  AW(0+k)=zgt(0)& wk
  'AW(0+k)=zgt(0)&wk
   
 rsgt.MoveNext	' 移到下一筆
 Wend
 
  'DOCTOR SexYKT, subtk
  'GGRNDSN BN3, YKGS

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %>
 
  <%  '計算天才壽年支
 ' SUB GGRNDSN(BN5, YKGS)
 FUNCTION GGRNDSN(BN3, YKGS)
 
   ZGG= Array("子","丑","寅","卯","辰","已","午","未","申","酉","戌","亥")
   ZGN= Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
      For n=0 to 11
        if BN3=ZGG(n) then
           'zgtk1=ZGN(n)
           zgtk1=n
          end if 
      Next   
     For m=0 to 11
        if YKGS=ZGG(m) then
           'zgtk2=ZGN(m)
           zgtk2=m
          end if 
     Next   
      zgtk3=zgtk1+zgtk2
      zgtk4=zgtk3 MOD 12
      ' zgtk=ZGG(zgtk4)
        ' 'GGRNDSN=ZG(YG)
      'zgtkn=ZGG(zgtk4)
      GGRNDSN=ZGG(zgtk4)
  'A(0+k)=zgt(0)
  ' B(0+k)=zgtk
 ' AW(0+k)=zgt(0)& wk
  ''AW(0+k)=zgt(0)&wk
  ' A(105)="天才壽"
  ' B(105)=zgtk
 ' AW(105)=zgt(0)& wk

' END SUB
      
 END FUNCTION 

  
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  %>
   
  <%  ''計算三台八座恩光天貴宮支
    'SUB TGRNDSN(TBNM, TDNUM)
   FUNCTION TGRNDSN(TBNM, TDNUM)
 
   ZGG= Array("子","丑","寅","卯","辰","已","午","未","申","酉","戌","亥")
   ZGN= Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
      For n=0 to 11
        if TBNM=ZGG(n) then
           'zmtk1=ZGN(n)
           zmtk1=n
           'zmtk2=TDNUM
         end if 
      Next   
    
       zmtk2=TDNUM
      zmtk3=zmtk1+zmtk2+36
      zmtk4=zmtk3 MOD 12
        'zmtk=ZGG(zmtk4)
        ''GGRNDSN=ZG(YG)
        TGRNDSN=ZGG(zmtk4)
 
 'A(0+k)=zmt(0)
  ' A(0+k)=三台
  ' B(0+k)=zmtk
 ' AW(0+k)=zmt(0)& wk
  
' END SUB
      
 END FUNCTION 

   
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %>   

<%
 ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From 月系星表"
Set rsmt = GetExcelRecordset( "Excel09.xls",  SQL)
If rsmt Is Nothing Then
    Response.Write "GetExcelRecordset 呼叫失敗!"
    Response.End
End If 
  ReDim zmt(16)
 ' zmt(0)="左右弼"
  'zmtk=" "
  wk=" "
  k=60


  ' Part II：輸出資料表的「內容」
  rsmt.MoveFirst		' 將目前資料錄移到第一筆
  While Not rsmt.EOF	' 判斷是否過了最後一筆
     k=k+1
     zmt(0)=rsmt(0)
    For i=1 to rsmt.Fields.Count-1
         'if Trim(rsmt(i).Name)=Trim(TNUM) then
      if Trim(rsmt(i).Name)=Trim("M"&TNUM) then
         'if rsmt(i).Name= "申" then
           'zmt(0)="左右弼"
           zmt(i)=rsmt(i)
           zmtk =rsmt(i)
           
      Select Case Trim(zmt(0))
        Case "三台"
          TDNUM1=DNUM-1
          zmtk=TGRNDSN(B(61), TDNUM1)
           ' DOCTOR SexYKT, subtk
         Case "八座" 
           TDNUM2=-(DNUM-1)
           zmtk=TGRNDSN(B(62), TDNUM2)
            ' GGRNDSN BN5, YKGS
         Case "恩光"
          TDNUM3=DNUM-2
          zmtk=TGRNDSN(B(54), TDNUM3)
           ' DOCTOR SexYKT, subtk
         Case "天貴" 
           TDNUM4=DNUM-2
           zmtk=TGRNDSN(B(55), TDNUM4)
            ' GGRNDSN BN5, YKGS
       End  Select      
      
      End If  
    Next
          
    ' rsmt.MoveNext	' 移到下一筆
  ' Wend 
 
  A(0+k)=zmt(0)
  B(0+k)=zmtk
  AW(0+k)=zmt(0)& wk
  'AW(0+k)=zmt(0)&wk
   
 rsmt.MoveNext	' 移到下一筆
 Wend
 
   ''DOCTOR SexYKT, subtk
   ''GGRNDSN BN3, YKGS
  ' TGRNDSN(TBNM, TDNUM)

%>


<%

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From 旬空表"
Set rssk = GetExcelRecordset( "Excel09.xls",  SQL)
If rssk Is Nothing Then
    Response.Write "GetExcelRecordset 呼叫失敗!"
    Response.End
End If 
  ReDim zsk(16)
 ' zsk(0)="旬空年支"
  'zskk=" "
  wk=" "
  k=105

 ' Part II：輸出資料表的「內容」
  rssk.MoveFirst		' 將目前資料錄移到第一筆
  While Not rssk.EOF	' 判斷是否過了最後一筆
   if  Trim(rssk(0))= Trim(YKG)  then
        ' zsk(0)=rssk(0)
       For i=1 to rssk.Fields.Count-1
        if Trim(rssk(i).Name)= Trim(YKK)  then
           'if rssk(i).Name= "申" then
           'zsk(i)="旬空宮位"
           zsk(i)=rssk(i)
           zskk =rssk(i)
         End If 
       Next
    ' rssk.Update
   ' End If  
   ' rssk.MoveNext	' 移到下一筆
 ' Wend  
    
   'A(0+k)=zsk(0)
   A(0+k)="旬空"
   B(0+k)=zskk
  'AW(0+k)=zsk(0)&wk
   AW(0+k)="旬空"&wk
   AKI(0+k)="F"&k
  End if  
 rssk.MoveNext	' 移到下一筆
 
 Wend
 
 A(106)="天傷"
   B(106)=B(12)
  'AW(0+k)=zsk(0)&wk
   AW(106)="天傷"&wk
   'A(12)="僕役宮"
 A(107)="天使"
   B(107)=B(10)
  'AW(0+k)=zsk(0)&wk
   AW(107)="天使"&wk
   ''A(10)="疾厄宮"

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 %>

<%

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From 四化星表"
Set rsk = GetExcelRecordset( "Excel09.xls",  SQL)
If rsk Is Nothing Then
    Response.Write "GetExcelRecordset 呼叫失敗!"
    Response.End
End If 
  ReDim zfk(16)
 ' zfk(0)="四化星宮"
  zfkk=" "
  wk=" "
  k=140

 ' Part II：輸出資料表的「內容」
  rsk.MoveFirst		' 將目前資料錄移到第一筆
  While Not rsk.EOF	' 判斷是否過了最後一筆
     k=k+1
   ' For j=1 to 11 
   ' IF rsk(0)= "十"  Then 
      'ReDim zfk(16)
     zfk(0)=rsk(0)
       For i=1 to rsk.Fields.Count-1
        if  rsk(i).Name= Trim(YKK)  then
         'if rsk(i).Name= z11k then
         'if rsk(i).Name= "申" then
           'zfk(0)="四化星宮"
           zfk(i)=rsk(i)
           zfkk =rsk(i)
         End If 
       Next
    ' rsk.Update
   ' End If  
   ' rsk.MoveNext	' 移到下一筆
 ' Wend  
    
   A(0+k)=zfk(0)
   B(0+k)=zfkk
   AW(0+k)=zfk(0)&wk
   AKI(0+k)="F"&k
       

 
 rsk.MoveNext	' 移到下一筆
 Wend
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
' Row = "<TR>"
 '   For i=0 to 100
 '     Row = Row & "<TD>" &  A(i)& B(i) & "</TD>"
     
    ''' For i=0 to rss.Fields.Count-1
      ''' Row = Row & "<TD>" & rss(i) & "</TD>"
      ''' Row = Row & "<TD>" & zf(i) & "</TD>"
  '   Next
  ' Response.Write Row & "</TR>"
  
 
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  

 ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From 大限表"
Set rsg = GetExcelRecordset( "Excel09.xls",  SQL)
If rsg Is Nothing Then
    Response.Write "GetExcelRecordset 呼叫失敗!"
    Response.End
End If 
  ReDim zfl(24)
 ' zfl(0)="大限表宮"
  zflk=" "
  wk=" "
  k=145

 ' Part II：輸出資料表的「內容」
  rsg.MoveFirst		' 將目前資料錄移到第一筆
  While Not rsg.EOF	' 判斷是否過了最後一筆
     k=k+1
   ' For j=1 to 11 
   ' IF rsg(0)= "十"  Then 
      'ReDim zfl(16)
     zfl(0)=rsg(0)
       For i=1 to rsg.Fields.Count-1
        if  rsg(i).Name= Trim(SexYK+z2k)  then
         ' if  rsg(i).Name= Trim(YKK)  then
         'if rsg(i).Name= z11k then
         'if rsg(i).Name= "申" then
           'zfl(0)="大限表宮"
           zfl(i)=rsg(i)
           zflk =rsg(i)
         End If 
       Next
    ' rsg.Update
   ' End If  
   ' rsg.MoveNext	' 移到下一筆
 ' Wend  
    
  ' A(0+k)=zfl(0)
  ' B(0+k)=zflk
  ' AW(0+k)=zfl(0)&wk

  A(0+k)=zflk
   B(0+k)=zfl(0)
  AW(0+k)=zflk&wk
   AKI(0+k)="G"&k
       

 rsg.MoveNext	' 移到下一筆
 Wend
 
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
' Row = "<TR>"
 '   For i=46 to 60
  '    Row = Row & "<TD>" &  A(i)& B(i) & "</TD>"
     
    '' For i=0 to rss.Fields.Count-1
      ''' Row = Row & "<TD>" & rss(i) & "</TD>"
      '' Row = Row & "<TD>" & zf(i) & "</TD>"
    ' Next
  ' Response.Write Row & "</TR>"
  
 
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From 小限表D"
Set rsgs = GetExcelRecordset( "Excel09.xls",  SQL)
If rsgs Is Nothing Then
    Response.Write "GetExcelRecordset 呼叫失敗!"
    Response.End
End If 
  ReDim zfms(24)
 ' zfms(0)="小限表宮"
  zfmsk=" "
  wk=" "
  k=160

 ' Part II：輸出資料表的「內容」
  rsgs.MoveFirst		' 將目前資料錄移到第一筆
  While Not rsgs.EOF	' 判斷是否過了最後一筆
     k=k+1
   ' For j=1 to 11 
   ' IF rsgs(0)= "十"  Then 
      'ReDim zfms(16)
     zfms(0)=rsgs(0)
       For i=1 to rsgs.Fields.Count-1
        if  rsgs(i).Name= Trim(Sex+YKG)  then
         ' if  rsgs(i).Name= Trim(YKK)  then
         'if rsgs(i).Name= z11k then
         'if rsgs(i).Name= "申" then
           'zfms(0)="小限表宮"
           zfms(i)=rsgs(i)
           zfmsk =rsgs(i)
         End If 
       Next
    ' rsgs.Update
   ' End If  
   ' rsgs.MoveNext	' 移到下一筆
 ' Wend  
    
   A(0+k)= zfms(0)
   B(0+k)=zfmsk
   AW(0+k)=zfms(0)&wk
    AKI(0+k)="SL"&k

  'A(0+k)=zfmsk
  'B(0+k)=zfms(0)
  'AW(0+k)=zfmsk&wk
  ' AKI(0+k)="G"&k
       

 
 rsgs.MoveNext	' 移到下一筆
 Wend
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
' Row = "<TR>"
 '   For i=46 to 75
 '     Row = Row & "<TD>" &  A(i)& B(i) & "</TD>"
     
 '   ' For i=0 to rss.Fields.Count-1
 '     ' Row = Row & "<TD>" & rss(i) & "</TD>"
 '     '' Row = Row & "<TD>" & zf(i) & "</TD>"
 '    Next
 '  Response.Write Row & "</TR>"
  
 
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'' ReDim R(13) 
   'ReDim KF(13)
   ' ReDim KFE(13)
  ReDim RKS(15,6)
 
  R= Array(" ","[子宮]","[丑宮]","[寅宮]","[卯宮]","[辰宮]","[已宮]","[午宮]","[未宮]","[申宮]","[酉宮]","[戌宮]","[亥宮]","[中宮]")

  RE= Array(" "," "," "," "," "," "," "," "," "," "," "," "," "," ")

 KFE= Array(" "," "," "," "," "," "," "," "," "," "," "," "," "," ")

  KF= Array(" ","[子宮]","[丑宮]","[寅宮]","[卯宮]","[辰宮]","[已宮]","[午宮]","[未宮]","[申宮]","[酉宮]","[戌宮]","[亥宮]","[中宮]")

 KS= Array("命宮","兄弟","夫妻","子女","財帛","疾厄","遷移","僕役","官祿","田宅","福德","父母","紫微","天機","太陽","武曲","天同","廉貞","天府","太陰","貪狼","巨門","天相","天梁","七殺","破軍","化祿","化權","化科","化忌"," ")
  KI= Array("K0","K1","K2","K3","K4","K5","K6","K7","K8","K9","K10","K11","S0","S1","S2","S3","S4","S5","T0","T1","T2","T3","T4","T5","T6","T7"," "," "," "," "," ")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ''替換宮名 
 For j=140 to 160
      BBG=B(j)
  For n=o to 200 
       AAk=A(n)  
     If BBG=AAK then
       B(j)=B(n)
     End if 
   Next
  Next  
  
 %>
   <%
 
    'SQL = "Select * From [B2:N12]"
 ''  cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPATH & ";" & DBpass
 ''   rs.Open "宮干表", cn, adOpenKeyset, adLockOptimistic
  
  SQL = "Select * From 宮干表"
 Set rsy4 = GetExcelRecordset( "Excel09.xls",  SQL)
  
 ' R = Array(" ", "[子宮]", "[丑宮]", "[寅宮]", "[卯宮]", "[辰宮]", "[已宮]", "[午宮]", "[未宮]", "[申宮]", "[酉宮]", "[戌宮]", "[亥宮]", "[中宮]")
  ReDim zt(13)
 zt(0) = "宮干"
   ' Part II：輸出資料表的「內容」
  rsy4.MoveFirst      ' 將目前資料錄移到第一筆
  While Not rsy4.EOF  ' 判斷是否過了最後一筆
    'If rsy4(0) = "庚" Then
    If rsy4(0) = Trim(YKK) Then
     'zt(0)="宮干"
       For i = 1 To rsy4.Fields.Count - 1
         zt(i) = rsy4(i) + rsy4(i).Name
        'rsy4(i) =  rsy4(i).Name+rsy4(i)
        R(i) = "[" + zt(i) + "宮]"
       Next
    End If
    
     rsy4.MoveNext    ' 移到下一筆
  Wend
  '' Text1.Text = R(5) + R(10)
  rsy4.Close
   '' cn.Close


 %>


<%

  ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From 流年諸星表"
Set rsy5 = GetExcelRecordset( "Excel09.xls",  SQL)
If rsy5 Is Nothing Then
    Response.Write "GetExcelRecordset 呼叫失敗!"
    Response.End
End If 
  ReDim zy5(16)
 ' zy5(0)="流年命宮諸星"
  zy5k=" "
  k=175
  wk=" "

  ' Part II：輸出資料表的「內容」
  rsy5.MoveFirst		' 將目前資料錄移到第一筆
  While Not rsy5.EOF	' 判斷是否過了最後一筆
         k=k+1
      zy5(0)=rsy5(0)
       For i=1 to rsy5.Fields.Count-1
         if rsy5(i).Name= Trim(LYG) then
          'if rsy5(i).Name= "寅" then
           'zy5(0)="流年命宮諸星"
           zy5(i)=rsy5(i)
           zy5k =rsy5(i)
          
        End If 
       Next
    ' rsy5.Update
   ' End If  
   ' rsy5.MoveNext	' 移到下一筆
 ' Wend  
  A(0+k)=zy5(0)
  B(0+k)=zy5k
  AW(0+k)=zy5(0)&wk
  AKI(0+k)="K"&k
   
  rsy5.MoveNext	' 移到下一筆
 Wend

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From 流年月君表"
Set rsy6 = GetExcelRecordset( "Excel09.xls",  SQL)
If rsy6 Is Nothing Then
    Response.Write "GetExcelRecordset 呼叫失敗!"
    Response.End
End If 
  ReDim zy6(16)
 ' zy6(0)="流年月君宮"
  zy6k=" "
  k=212
  wk=" "

  ' Part II：輸出資料表的「內容」
  rsy6.MoveFirst		' 將目前資料錄移到第一筆
  While Not rsy6.EOF	' 判斷是否過了最後一筆
       
    If TRIM(rsy6(0))=Trim(TNUM) Then 
         'zy6(0)=rsy6(0)
       For i=1 to rsy6.Fields.Count-1
         if rsy6(i).Name= Trim(LYG) then
          'if rsy6(i).Name= "寅" then
           'zy6(0)="身命宮"
           zy6(i)=rsy6(i)
          zy66k =rsy6(i)
         zy6k= LGRNDSN(zy66k, HNUM)
    
        End If 
       Next
    ' rsy6.Update
   End If  
   ' rsy6.MoveNext	' 移到下一筆
 ' Wend  
  'A(0+k)=zy6(0)
  A(212)="正月君"
  B(212)=zy6k
 ' AW(0+k)="正月君"zy6(0)&wk
  AW(0+k)="正月君"&wk
  AKI(0+k)="K"&k
   
  rsy6.MoveNext	' 移到下一筆
 Wend
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
%>
  
  <%  '計算流年月君
 ' SUB LGRNDSN(MONG, HNUMR)
 FUNCTION LGRNDSN(MONG, HNUMR)
 
   ZGG= Array("子","丑","寅","卯","辰","已","午","未","申","酉","戌","亥")
   ZGN= Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
      For n=0 to 11
        if MONG=ZGG(n) then
           'zgtk1=ZGN(n)
           zgtk1=n
          end if 
      Next   
     For m=0 to 11
        if HNUMR=ZGG(m) then
           'zgtk2=ZGN(m)
           zgtk2=m
          end if 
     Next   
      zgtk3=zgtk1+zgtk2
      zgtk4=zgtk3 MOD 12
      ' zgtk=ZGG(zgtk4)
        
      LGRNDSN=ZGG(zgtk4)
  
' END SUB
      
 END FUNCTION 

  %>    
 
 <%
' Row = "<TR>"
'    For i=90 to 220
 '     Row = Row & "<TD>" & AW(i)& B(i) & "</TD>"
     
    ''' For i=0 to rss.Fields.Count-1
      ''' Row = Row & "<TD>" & rss(i) & "</TD>"
     ' '' Row = Row & "<TD>" & zf(i) & "</TD>"
 '   Next
 '  Response.Write Row & "</TR>"

 
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  For j=0 to 220
      BBK=Trim(B(j))
   Select Case BBK
      Case "子"
        R(1)=Trim(R(1))+","& Trim(AW(j))
        RE(1)=Trim(RE(1))+Trim(AW(j))
         KFE(1)=KFE(1)+AKI(j)
          Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(1,1)=Trim(RKS(1,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(1,2)=Trim(RKS(1,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(1,3)=Trim(RKS(1,1))+Trim(A(j))
          Case 41, 42, 43, 44, 45,46, 47
              RKS(1,4)=Trim(RKS(1,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(1,5)=Trim(RKS(1,1))+Trim(A(j))
           Case Else
              RKS(1,6)=Trim(RKS(1,1))+Trim(A(j))
         End  Select      

    

       Case "丑"
        R(2)=R(2)+","&AW(j)  
        RE(2)=Trim(RE(2))+Trim(AW(j))
         KFE(2)=KFE(2)+AKI(j)
          Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(2,1)=Trim(RKS(2,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(2,2)=Trim(RKS(2,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(2,3)=Trim(RKS(2,1))+Trim(A(j))
           Case 41, 42, 43, 44, 45,46, 47
              RKS(2,4)=Trim(RKS(2,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(2,5)=Trim(RKS(2,1))+Trim(A(j))
 
           Case Else
              RKS(2,6)=Trim(RKS(2,1))+Trim(A(j))
         End  Select      


       Case "寅"
        R(3)=R(3)+","&AW(j) 
         RE(3)=Trim(RE(3))+Trim(AW(j)) 
          KFE(3)=KFE(3)+AKI(j)
           Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(3,1)=Trim(RKS(3,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(3,2)=Trim(RKS(3,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(3,3)=Trim(RKS(3,1))+Trim(A(j))
            Case 41, 42, 43, 44, 45,46, 47
              RKS(3,4)=Trim(RKS(3,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(3,5)=Trim(RKS(3,1))+Trim(A(j))
           Case Else
              RKS(3,6)=Trim(RKS(3,1))+Trim(A(j))
         End  Select      

             
      Case "卯"
        R(4)=R(4)+","&AW(j)
        RE(4)=Trim(RE(4))+Trim(AW(j)) 
         KFE(4)=KFE(4)+AKI(j) 
          Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(4,1)=Trim(RKS(4,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(4,2)=Trim(RKS(4,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(4,3)=Trim(RKS(4,1))+Trim(A(j))
          Case 41, 42, 43, 44, 45,46, 47
              RKS(4,4)=Trim(RKS(4,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(4,5)=Trim(RKS(4,1))+Trim(A(j))
 
           Case Else
              RKS(4,6)=Trim(RKS(4,1))+Trim(A(j))
         End  Select      


      Case "辰"
        R(5)=R(5)+","&AW(j)
        RE(5)=Trim(RE(5))+Trim(AW(j)) 
         KFE(5)=KFE(5)+AKI(j)
           Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(5,1)=Trim(RKS(5,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(5,2)=Trim(RKS(5,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(5,3)=Trim(RKS(5,1))+Trim(A(j))
           Case 41, 42, 43, 44, 45,46, 47
              RKS(5,4)=Trim(RKS(5,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(5,5)=Trim(RKS(5,1))+Trim(A(j))
           Case Else
              RKS(5,6)=Trim(RKS(5,1))+Trim(A(j))
         End  Select      


     Case "已"
        R(6)=R(6)+","&AW(j) 
        RE(6)=Trim(RE(6))+Trim(AW(j))
         KFE(6)=KFE(6)+AKI(j)
           Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(6,1)=Trim(RKS(6,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(6,2)=Trim(RKS(6,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(6,3)=Trim(RKS(6,1))+Trim(A(j))
            Case 41, 42, 43, 44, 45,46, 47
              RKS(6,4)=Trim(RKS(6,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(6,5)=Trim(RKS(6,1))+Trim(A(j))
            Case Else
              RKS(6,6)=Trim(RKS(6,1))+Trim(A(j))
         End  Select      

    
      Case "午"
        R(7)=R(7)+","&AW(j) 
        RE(7)=Trim(RE(7))+Trim(AW(j))
         KFE(7)=KFE(7)+AKI(j)
          Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(7,1)=Trim(RKS(7,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(7,2)=Trim(RKS(7,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(7,3)=Trim(RKS(7,1))+Trim(A(j))
           Case 41, 42, 43, 44, 45,46, 47
              RKS(7,4)=Trim(RKS(7,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(7,5)=Trim(RKS(7,1))+Trim(A(j))
           Case Else
              RKS(7,6)=Trim(RKS(7,1))+Trim(A(j))
          End  Select      
       

       Case "未"
        R(8)=R(8)+","&AW(j)
        RE(8)=Trim(RE(8))+Trim(AW(j))
         KFE(8)=KFE(8)+AKI(j)
         Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(8,1)=Trim(RKS(8,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(8,2)=Trim(RKS(8,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(8,3)=Trim(RKS(8,1))+Trim(A(j))
            Case 41, 42, 43, 44, 45,46, 47
              RKS(8,4)=Trim(RKS(8,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(8,5)=Trim(RKS(8,1))+Trim(A(j))
           Case Else
              RKS(8,6)=Trim(RKS(8,1))+Trim(A(j))
         End  Select      

      Case "申"
        R(9)=R(9)+","&AW(j)
        RE(9)=Trim(RE(9))+Trim(AW(j)) 
          KFE(9)=KFE(9)+AKI(j)
            Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(9,1)=Trim(RKS(9,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(9,2)=Trim(RKS(9,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(9,3)=Trim(RKS(9,1))+Trim(A(j))
            Case 41, 42, 43, 44, 45,46, 47
              RKS(9,4)=Trim(RKS(9,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(9,5)=Trim(RKS(9,1))+Trim(A(j))
           Case Else
              RKS(9,6)=Trim(RKS(9,1))+Trim(A(j))
         End  Select      
           

       Case "酉"
        R(10)=R(10)+","&AW(j) 
        RE(10)=Trim(RE(10))+Trim(AW(j)) 
         KFE(10)=KFE(10)+AKI(j)
          Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(10,1)=Trim(RKS(10,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(10,2)=Trim(RKS(10,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(10,3)=Trim(RKS(10,1))+Trim(A(j))
          Case 41, 42, 43, 44, 45,46, 47
              RKS(10,4)=Trim(RKS(10,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(10,5)=Trim(RKS(10,1))+Trim(A(j))
           Case Else
              RKS(10,6)=Trim(RKS(10,1))+Trim(A(j))
          End  Select      


       Case "戌"
        R(11)=R(11)+","&AW(j)
        RE(11)=Trim(RE(11))+Trim(A(j))
         KFE(11)=KFE(11)+AKI(j)
          Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(11,1)=Trim(RKS(11,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(11,2)=Trim(RKS(11,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(11,3)=Trim(RKS(11,1))+Trim(A(j))
            Case 41, 42, 43, 44, 45,46, 47
              RKS(11,4)=Trim(RKS(11,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(11,5)=Trim(RKS(11,1))+Trim(A(j))
           Case Else
              RKS(11,6)=Trim(RKS(11,1))+Trim(A(j))
         End  Select      


      Case "亥"
       R(12)=R(12)+","&AW(j)
       RE(12)=Trim(RE(12))+Trim(AW(j)) 
        KFE(12)=KFE(12)+AKI(j)
         Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(12,1)=Trim(RKS(12,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(12,2)=Trim(RKS(12,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(12,3)=Trim(RKS(12,1))+Trim(A(j))
            Case 41, 42, 43, 44, 45,46, 47
              RKS(12,4)=Trim(RKS(12,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(12,5)=Trim(RKS(12,1))+Trim(A(j))
           Case Else
            RKS(12,6)=Trim(RKS(12,1))+Trim(A(j))
          End  Select      

       Case Else
        R(13)=R(13)&AW(j)
        RE(13)=Trim(RE(13))+Trim(A(j)) 
         KFE(13)=KFE(13)+AKI(j)
         Select Case j
           Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
              RKS(13,1)=Trim(RKS(13,1))+Trim(A(j))
           Case 21, 22, 23, 24, 25, 26, 27, 28, 29, 30
              RKS(13,2)=Trim(RKS(13,1))+Trim(A(j))
            Case 31, 32, 33, 34, 35, 36, 37, 38, 39, 40
              RKS(13,3)=Trim(RKS(13,1))+Trim(A(j))
           Case 41, 42, 43, 44, 45,46, 47
              RKS(13,4)=Trim(RKS(13,1))+Trim(A(j))
          Case 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
              RKS(13,5)=Trim(RKS(13,1))+Trim(A(j))
           Case Else
              RKS(13,6)=Trim(RKS(13,1))+Trim(A(j))
         End  Select      
             
     End  Select      
  Next      

%>
 
</TABLE></CENTER>
  
 <CENTER>
<TABLE Border=1 Width=80%  BGCOLOR=#FFFF00>--->
 <TR><TD width=230></TD><TD width=230></TD><TD width =230></TD><TD width =230></TD></TR>

 <% 

Row1 = "<TR >" & "<TD width=230>" & Trim(R(6)) & "</TD>"& "<TD width=230>" & Trim(R(7)) & "</TD>"& "<TD width =230>" & Trim(R(8)) & "</TD>"& "<TD width =230>" & Trim(R(9)) & "</TD>"& "</TR>"
Row2 = "<TR>" & "<TD width=230>" & Trim(R(5)) & "</TD>"& "<TD width=230  RowSpan=2 ColSpan=2>"& Trim(R(13))& SECTM&TNUM&DNUM&HNUM&",命主:"&MYL&",身主:"&MYB& "</TD>"& "<TD width =230>" & Trim(R(10)) & "</TD>"& "</TR>"
Row3 = "<TR>" & "<TD width=230>" & Trim(R(4)) & "</TD>"&  "<TD width=230>" & Trim(R(11)) & "</TD>"& "</TR>"
Row4 = "<TR>" & "<TD width=230>" & Trim(R(3)) & "</TD>"& "<TD width=230>" & Trim(R(2)) & "</TD>"& "<TD width =230>" & Trim(R(1)) & "</TD>"& "<TD width =230>" & Trim(R(12)) & "</TD>"& "</TR>"
RowT=Row1+Row2+Row3+Row4
 
 Response.Write RowT 
' Response.Write A(5)+B(5)+DNUM+HNUM 
 
 
  %>
  
 </TABLE></CENTER>
 
 
  <%
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '''SQL = "Select * From [A2:B100]"
 ' SQL = "Select * From 評論表2"
   '''SQL = "Select * From K0_1"
 'Set rls = GetExcelRecordset( "Excel12.xls",  SQL)
 %>

  <!-- <CENTER>
<TABLE Border=1 BGCOLOR=#FFFF00>    
 <TR ><TD width=720></TD></TR>-->

 
 <!--</TABLE></CENTER>-->

 <%
   
   ''   Row = "<TR>"
 ' Row = "<TR>"
    ''   For i=1 to 75
     '''   Row = Row & "<TD>" & AW(i)& B(i) & "</TD>"
 '   For i=0 to 13
 '       Row = Row & "<TD>" & R(i) & "</TD>"
     ''' Row = Row & "<TD>" & KFE(i) & "</TD>"
          
 '  Next
 '  Response.Write Row & "</TR>"

 
 %>
<!-- </TABLE></CENTER>-->

<!-- <TABLE BORDER=1>
<TR BGCOLOR=#00FFFF>-->

 <%
   ''' Row3 = "<TR>"
 '   For i=1 to 13
 '    For J=1 to  6
      '''  For J=1 to  5
   
 '   Response.Write "<TD>" & RKS(i,j) & "</TD>"

 '     Next  
 '  Next
 
%>
 <!--</TABLE>--> 

  <CENTER><H2><FONT color=#FF0000>若需要命 盤 評 論 表,請加入<u>免費會員</u>,  謝謝</FONT><HR></H2>
  <!--<CENTER><h2><a href="http://class.ruten.com.tw/user/index00.php?s=tang1206">回拍賣場</a></h2>-->
  <CENTER><h2 align="left"><a href="http://61.222.248.199/HSU-fundb/Login-2r.asp" ><u>加入會員</u></a></h2>


</BODY>
<!-- <script>
    document.onselectionchange=__OnSelectionChange;
       var running=false;
     function __OnSelectionChange()
       { 
       if (running==true) return;
          running=true;
       document.selection.empty();
       running=false;       
        }
  </script>--> 

</HTML>



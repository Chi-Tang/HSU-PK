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
 SQL = "Select * From �R�c��"
Set rs3 = GetExcelRecordset( "Excel08.xls",  SQL)
If rs3 Is Nothing Then
    Response.Write "GetExcelRecordset �I�s����!"
    Response.End
End If 
  ReDim z3(16)
  z3(0)="�R�c"
  'z3k=" "
   wk=" "
   ZLG= Array("�l","��","�G","�f","��","�w","��","��","��","��","��","��")
   ZLN= Array("�g�T","����","�S�s","�妱","�G�s","�Z��","�}�x","�Z��","�G�s","�妱","�S�s","����")
   ZBN= Array("���P","�Ѭ�","�ѱ�","�ѦP","���","�Ѿ�","���P","�Ѭ�","�ѱ�","�ѦP","���","�Ѿ�")

 %>
<HTML>
<BODY bgcolor="#FFFFFF">
  <CENTER><H2>�� �L �� �� �R �L �P �� �� �� �p �U<HR></H2>
<TABLE BORDER=1>
<TR BGCOLOR=#00FFFF>
 </TR>
 
<%
  ' Part II�G��X��ƪ��u���e�v
  rs3.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
  While Not rs3.EOF	' �P�_�O�_�L�F�̫�@��
    IF Trim(rs3(0))= Trim(TNUM)  Then
   'IF rs3(0)= "9"  Then 
      'ReDim z3(16)
          ' z3(0)=rs3(0)    
       For i=1 to rs3.Fields.Count-1
        if Trim(rs3(i).Name)= Trim(HNUM) then
        'if rs3(i).Name= "��" then
           z3(i)=rs3(i)
           z3k =rs3(i)
          ' z3(0)="(��)�R�c"+z3k
          ' z30k="(��)�R�c"
          ' z3(i) =  rs3(i)+rs3(i).Name
         ''rs3(i) =rs3(i).Name+rs3(i)
        ''' Exit For
        End If 
       Next
    ' rs3.Update
    End If  
     rs3.MoveNext	' ����U�@��
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
 SQL = "Select * From ���c��"
Set rsb1 = GetExcelRecordset( "Excel08.xls",  SQL)
If rsb1 Is Nothing Then
    Response.Write "GetExcelRecordset �I�s����!"
    Response.End
End If 
  ReDim zb1(16)
  zb1(0)="���c"
  'zb1k=" "
   wk=" "
   
  ' Part II�G��X��ƪ��u���e�v
  rsb1.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
  While Not rsb1.EOF	' �P�_�O�_�L�F�̫�@��
    IF Trim(rsb1(0))= Trim(TNUM)  Then
   'IF rsb1(0)= "9"  Then 
       For i=1 to rsb1.Fields.Count-1
        if Trim(rsb1(i).Name)= Trim(HNUM) then
        'if rsb1(i).Name= "��" then
           zb1(i)=rsb1(i)
           zb1k =rsb1(i)
         End If 
       Next
     'rsb1.Update
    End If  
     rsb1.MoveNext	' ����U�@��
  Wend
 
 A(3)=zb1(0)
 B(3)=zb1k
 AW(3)=zb1(0)&wk
 AKI(3)="K3"
 
 %>
<%'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From �l�c��"
Set rs5 = GetExcelRecordset( "Excel08.xls",  SQL)
If rs5 Is Nothing Then
    Response.Write "GetExcelRecordset �I�s����!"
    Response.End
End If 
  ReDim z5(16)
 ' z5(0)="���R�c"
  z5k=" "
  k=5
  wk=" "

%>
<%
  '' Part I�G��X�u���Y�W�١v
  'For i=0 to rs5.Fields.Count-1
  '  Response.Write "<TD>" & rs5(i).Name & "</TD>"
 ' Next
%>

<%
  ' Part II�G��X��ƪ��u���e�v
  rs5.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
  While Not rs5.EOF	' �P�_�O�_�L�F�̫�@��
         k=k+1
   ' For j=1 to 11 
    ' IF rs5(0)= "�Q"  Then 
      'ReDim z5(16)
      z5(0)=rs5(0)
       For i=1 to rs5.Fields.Count-1
         if rs5(i).Name= z3k then
          'if rs5(i).Name= "�G" then
           'z5(0)="���R�c"
           z5(i)=rs5(i)
           z5k =rs5(i)
          ' z5(i) =  rs5(i)+rs5(i).Name
         ''rs5(i) =rs5(i).Name+rs5(i)
        ''' Exit For
        End If 
       Next
    ' rs5.Update
   ' End If  
   ' rs5.MoveNext	' ����U�@��
 ' Wend  
  A(0+k)=z5(0)
  B(0+k)=z5k
  AW(0+k)=z5(0)&wk
  AKI(0+k)="K"&k
 
    
  rs5.MoveNext	' ����U�@��
 Wend

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
  'SQL = "Select * From [B2:N12]"
 SQL = "Select * From ���槽"
Set rs2 = GetExcelRecordset( "Excel08.xls",  SQL)
If rs2 Is Nothing Then
    Response.Write "GetExcelRecordset �I�s����!"
    Response.End
End If 
  ReDim z2(16)
  z2(0)="(��)�R�c"
 '' z2k=" "
  wk=" "

%>


<%
  ' Part II�G��X��ƪ��u���e�v
  rs2.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
  While Not rs2.EOF	' �P�_�O�_�L�F�̫�@��
    IF Trim(rs2(0))= Trim(YKK)  Then 
   
    '  IF rs2(0)= "��"  Then 
      'ReDim z2(16)
         ' z2(0)=rs2(0)
       For i=1 to rs2.Fields.Count-1
          if rs2(i).Name= z3k then
          ' if rs2(i).Name= "�G" then
           'z2(0)="(��)�R�c"
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
     rs2.MoveNext	' ����U�@��
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
 SQL = "Select * From ���L�P"
Set rs1 = GetExcelRecordset( "Excel08.xls",  SQL)
If rs1 Is Nothing Then
    Response.Write "GetExcelRecordset �I�s����!"
    Response.End
End If 
  ReDim z1(16)
  z1(0)="���L�P"
  z1k=" "
  z11k=" "
  ZB= Array("�l","��","�G","�f","��","�w","��","��","��","��","��","��")
  
  ZF= Array("��","�f","�G","��","�l","��","��","��","��","��","��","�w")

%>

<%
  ' Part II�G��X��ƪ��u���e�v
  rs1.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
  While Not rs1.EOF	' �P�_�O�_�L�F�̫�@��
    IF Trim(rs1(0))= Trim(DNUM)  Then 
    'IF rs1(0)= "8"  Then 
      'ReDim z1(16)
     
       For i=1 to rs1.Fields.Count-1
        if rs1(i).Name= z2k then
         'if rs1(i).Name= "��T��" then
           '''z1(0)="(���L)�ѩ��P"
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
     rs1.MoveNext	' ����U�@��
  Wend 
%> 

<%
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From �ҬP�L"
Set rs = GetExcelRecordset( "Excel08.xls",  SQL)
If rs Is Nothing Then
    Response.Write "GetExcelRecordset �I�s����!"
    Response.End
End If 
  ReDim z(16)
 ' z(0)="(���L)�ѩ��c"
  zk=" "
  k=20
%>


<%
  ' Part II�G��X��ƪ��u���e�v
  rs.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
  While Not rs.EOF	' �P�_�O�_�L�F�̫�@��
      k=k+1
   ' For j=1 to 11 
   ' IF rs(0)= "�Q"  Then 
          z(0)=rs(0)
         ' zzo=rs(0)
       For i=1 to rs.Fields.Count-1
        if rs(i).Name= z1k then
         'if rs(i).Name= "��" then
           'z(0)="(���L)�ѩ��c"
           z(i)=rs(i)
           zk =rs(i)
         End If 
       Next
    ' rs.Update
   ' End If  
   ' rs.MoveNext	' ����U�@��
 ' Wend 
 
 
  ''SQL = "Select * From �ҬP����"([B2:N12]") 
   ' WKW=WRNDSN(zzo, zk)
   wkw=WRNDSN(z(0), zk)
  
  A(0+k)=z(0)
  B(0+k)=zk
  AW(0+k)=z(0)&wkw
    Select Case Trim(wkw)
     Case "�q","��"
        AKI(0+k)="S"&k&"A"
     Case "�o","��"
        AKI(0+k)="S"&k&"B"
    Case Else
        AKI(0+k)="S"&k&"C"
    End  Select      
 
 rs.MoveNext	' ����U�@��
 Wend
  
%>

<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
FUNCTION WRNDSN(zzo, zk)
'' WRNDSN=WK  END FUNCTION 
 SQL = "Select * From �ҬP����"
Set rsd = GetExcelRecordset( "Excel08.xls",  SQL)
 'If rsd Is Nothing Then
 '   Response.Write "GetExcelRecordset �I�s����!"
  '  Response.End
 'End If 
  ReDim w(16)
  w(0)="����"  
  wk=" "

  ' Part II�G��X��ƪ��u���e�v
  rsd.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
  While Not rsd.EOF	' �P�_�O�_�L�F�̫�@��
    IF rsd(0)= zzo  Then
    'IF rsd(0)= z(0)  Then
    'IF rsd(0)= "(���L)�ѩ�"  Then 
      'ReDim z(16)
     
       For i=1 to rsd.Fields.Count-1
        if rsd(i).Name=  zk then
          'if rsd(i).Name= "�G" then
          ' w(0)="����"
           w(i)=rsd(i)
           wk =rsd(i)
         End If 
       Next
     rsd.Update
    End If  
     rsd.MoveNext	' ����U�@��
  Wend 
 WRNDSN=wk  
END FUNCTION 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

SUB  WRNDSNft(zzo, zftk)

  
  ''SQL = "Select * From [B2:N12]"
 SQL = "Select * From �ҬP����"
Set rsdt = GetExcelRecordset( "Excel09.xls",  SQL)
 'If rsdt Is Nothing Then
 '   Response.Write "GetExcelRecordset �I�s����!"
  '  Response.End
 'End If 
  ReDim w(16)
  w(0)="����"
  wk=" "

  ' Part II�G��X��ƪ��u���e�v
  rsdt.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
  While Not rsdt.EOF	' �P�_�O�_�L�F�̫�@��
    IF rsd(0)= zzo  Then
    'IF rsdt(0)= zft(0)  Then
    'IF rsdt(0)= "(���L)�ѩ�"  Then 
      'ReDim zft(16)
     
       For i=1 to rsdt.Fields.Count-1
        if rsdt(i).Name=  zftk then
          'if rsdt(i).Name= "�G" then
          ' w(0)="����"
           w(i)=rsdt(i)
           wk =rsdt(i)
         End If 
       Next
     rsdt.Update
    End If  
     rsdt.MoveNext	' ����U�@��
  Wend 
   A(0+k)=zft(0)
  B(0+k)=zftk
  AW(0+k)=zft(0)& wk
  'AW(0+k)=zft(0)&wk
 
  AW(0+k)=zft(0)&wk
    Select Case Trim(wk)
     Case "�q","��"
        AKI(0+k)="S"&k&"A"
     Case "�o","��"
        AKI(0+k)="S"&k&"B"
    Case Else
        AKI(0+k)="S"&k&"C"
    End  Select      
 
  
  END SUB
%>

<%
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
 ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From �ҬP�L2"
Set rsf = GetExcelRecordset( "Excel08.xls",  SQL)
If rsf Is Nothing Then
    Response.Write "GetExcelRecordset �I�s����!"
    Response.End
End If 
  ReDim zf(16)
 ' zf(0)="(���L)�ѩ��c"
  zfk=" "
  k=30
%>


<%
  ' Part II�G��X��ƪ��u���e�v
  rsf.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
  While Not rsf.EOF	' �P�_�O�_�L�F�̫�@��
     k=k+1
   ' For j=1 to 11 
   ' IF rsf(0)= "�Q"  Then 
      'ReDim zf(16)
     zf(0)=rsf(0)
       For i=1 to rsf.Fields.Count-1
        if rsf(i).Name= z11k then
         'if rsf(i).Name= "��" then
           'zf(0)="(���L)�ѩ��c"
           zf(i)=rsf(i)
           zfk =rsf(i)
         End If 
       Next
    ' rsf.Update
   ' End If  
   ' rsf.MoveNext	' ����U�@��
 ' Wend 
  ''SQL = "Select * From �ҬP����"([B2:N12]") 
   ' WKW=WRNDSN(zzo, zfk)
   wkw=WRNDSN(zf(0), zfk)
  
  A(0+k)=zf(0)
  B(0+k)=zfk
  AW(0+k)=zf(0)&wkw
 
  Select Case Trim(wkw)
     Case "�q","��"
        AKI(0+k)="T"&k&"A"
     Case "�o","��"
        AKI(0+k)="T"&k&"B"
    Case Else
        AKI(0+k)="T"&k7&"C"
    End  Select      
 
  
 rsf.MoveNext	' ����U�@��
 Wend
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Select Case Trim(YKK)
     Case "��","��","��","��","��"
         SexYk="��"&Sex
           Select Case Trim(Sex)
             Case "�k"
                 SexYKT="S"
             Case "�k"
                 SexYKT="R"
            ' Case Else
              ' SexYKT="N""
             End  Select   
         
     Case "�A","�B","�v","��","��"
         SexYk="��"&Sex
            Select Case Trim(Sex)
             Case "�k"
                 SexYKT="R"
             Case "�k"
                 SexYKT="S"
            ' Case Else
              ' SexYKT="N""
             End  Select 
     Case Else
         SexYk="��"&Sex
    End  Select   
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From �ҬP�L3"
Set rsft = GetExcelRecordset( "Excel09.xls",  SQL)
If rsft Is Nothing Then
    Response.Write "GetExcelRecordset �I�s����!"
    Response.End
End If 
  ReDim zft(16)
 ' zft(0)="�S�s�P"
  zftk=" "
  k=40


  ' Part II�G��X��ƪ��u���e�v
  rsft.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
  While Not rsft.EOF	' �P�_�O�_�L�F�̫�@��
     k=k+1
   ' For j=1 to 11 
   ' IF rsft(0)= "�Q"  Then 
        zft(0)=rsft(0)
       For i=1 to rsft.Fields.Count-1
        if rsft(i).Name= Trim(YKK) then
         'if rsft(i).Name= "��" then
           'zft(0)="�S�s�P"
           zft(i)=rsft(i)
           zftk =rsft(i)
           
          IF Trim(rsft(0))= "�S�s"  Then 
             subtk=zftk 
              
             ' DOCTOR SexYKT, subtk
           End If
       End If  
     Next
          
    ' rsft.MoveNext	' ����U�@��
  ' Wend 
  ''SQL = "Select * From �ҬP����"([B2:N12]") 
   ' WKW=WRNDSN(zzo, zftk)
   wkw=WRNDSN(zft(0), zftk)
    
  A(0+k)=zft(0)
  B(0+k)=zftk
  AW(0+k)=zft(0)& wkw
  'AW(0+k)=zft(0)&wkw
   
 rsft.MoveNext	' ����U�@��
 Wend
 
  DOCTOR SexYKT, subtk

   %>
 
 <%
 
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 SUB DOCTOR(SexYKT, zftk)
 
 ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From �դh��"
Set rspt = GetExcelRecordset( "Excel09.xls",  SQL)
If rspt Is Nothing Then
    Response.Write "GetExcelRecordset �I�s����!"
    Response.End
End If 
  ReDim zpt(24)
 ' zpt(0)="�S�s�P"
   wk=" "
  k=75


  ' Part II�G��X��ƪ��u���e�v
  rspt.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
  While Not rspt.EOF	' �P�_�O�_�L�F�̫�@��
     k=k+1
   ' For j=1 to 11 
   ' IF rspt(0)= "�Q"  Then 
      'ReDim zpt(24)
     zpt(0)=rspt(0)
       For i=1 to rspt.Fields.Count-1
        if rspt(i).Name= Trim(SexYKT&zftk) then
         'if rspt(i).Name= "��" then
           'zpt(0)="�S�s�P"
           zpt(i)=rspt(i)
           zptk =rspt(i)
         End If 
       
       Next
          
    ' rspt.MoveNext	' ����U�@��
  ' Wend  
 
  A(0+k)=zpt(0)
  B(0+k)=zptk
  AW(0+k)=zpt(0)&wk
  'AW(0+k)=zpt(0)&wk

    rspt.MoveNext	' ����U�@��
   Wend
 
 END SUB
 
  %>
 
 
 <% 
 'SQL = "Select * From [B2:N12]"
 SQL = "Select * From ���P��"
Set rs6 = GetExcelRecordset( "Excel09.xls",  SQL)
If rs6 Is Nothing Then
    Response.Write "GetExcelRecordset �I�s����!"
    Response.End
End If 
  ReDim z6(16)
  z6(0)="���P"
  z6k=" "
  k=51

  ' Part II�G��X��ƪ��u���e�v
  rs6.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
  While Not rs6.EOF	' �P�_�O�_�L�F�̫�@��
       '' z6(0)=rs6(0)
       z6(0)="���P"
   IF Trim(rs6(0))= Trim(YKG)  Then 
    'IF rs6(0)= "8"  Then 
          
       For i=1 to rs6.Fields.Count-1
        if rs6(i).Name= Trim(HNUM) then
         'if rs6(i).Name= "��T��" then
           
           z6(i)=rs6(i)
           z6k =rs6(i)
         
        End If 
       Next
     rs6.Update
    End If  
     rs6.MoveNext	' ����U�@��
  Wend 
%> 

<%
    
  ''SQL = "Select * From �ҬP����"([B2:N12]") 
   ' WKW=WRNDSN(zzo, zftk)
   wkw=WRNDSN(z6(0), z6k)
    
  A(0+k)=z6(0)
  B(0+k)=z6k
  AW(0+k)=z6(0)&wkw
      
%>


<% 
 'SQL = "Select * From [B2:N12]"
 SQL = "Select * From �a�P��"
Set rs7 = GetExcelRecordset( "Excel09.xls",  SQL)
If rs7 Is Nothing Then
    Response.Write "GetExcelRecordset �I�s����!"
    Response.End
End If 
  ReDim z7(16)
  z7(0)="�a�P"
  z7k=" "
  k=52

  ' Part II�G��X��ƪ��u���e�v
  rs7.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
  While Not rs7.EOF	' �P�_�O�_�L�F�̫�@��
       '' z7(0)=rs7(0)
       z7(0)="�a�P"
   IF Trim(rs7(0))= Trim(YKG)  Then 
    'IF rs7(0)= "8"  Then 
          
       For i=1 to rs7.Fields.Count-1
        if rs7(i).Name= Trim(HNUM) then
         'if rs7(i).Name= "��T��" then
           
           z7(i)=rs7(i)
           z7k =rs7(i)
         
        End If 
       Next
     rs7.Update
    End If  
     rs7.MoveNext	' ����U�@��
  Wend 
%> 

<%
    
  ''SQL = "Select * From �ҬP����"([B2:N12]") 
   ' WKW=WRNDSN(zzo, zftk)
   wkw=WRNDSN(z7(0), z7k)
    
  A(0+k)=z7(0)
  B(0+k)=z7k
  AW(0+k)=z7(0)&wkw
      
%>


<%
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From �ҬP�L4"
Set rsf7 = GetExcelRecordset( "Excel09.xls",  SQL)
If rsf7 Is Nothing Then
    Response.Write "GetExcelRecordset �I�s����!"
    Response.End
End If 
  ReDim zf7(16)
 ' zf7(0)="���(��)�P"
  zf7k=" "
  k=53

  ' Part II�G��X��ƪ��u���e�v
  rsf7.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
  While Not rsf7.EOF	' �P�_�O�_�L�F�̫�@��
     k=k+1
     zf7(0)=rsf7(0)
       For i=1 to rsf7.Fields.Count-1
        if rsf7(i).Name= Trim(HNUM) then
         'if rsf7(i).Name= "��" then
           'zf7(0)="(���L)�ѩ��c"
           zf7(i)=rsf7(i)
           zf7k =rsf7(i)
         End If 
       Next
    ' rsf7.Update
   ' End If  
   ' rsf7.MoveNext	' ����U�@��
 ' Wend  
   ''SQL = "Select * From �ҬP����"([B2:N12]") 
   ' WKW=WRNDSN(zzo, zf7k)
   wkw=WRNDSN(zf7(0), zf7k)
    
  A(0+k)=zf7(0)
  B(0+k)=zf7k
  AW(0+k)=zf7(0)&wkw
 
 rsf7.MoveNext	' ����U�@��
 Wend
 
  %>
  
  <%

 
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From �~��P��"
Set rsgt = GetExcelRecordset( "Excel09.xls",  SQL)
If rsgt Is Nothing Then
    Response.Write "GetExcelRecordset �I�s����!"
    Response.End
End If 
  ReDim zgt(16)
 ' zgt(0)="�Ѥ~��"
  zgtk=" "
  wk=" "
  k=90


  ' Part II�G��X��ƪ��u���e�v
  rsgt.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
  While Not rsgt.EOF	' �P�_�O�_�L�F�̫�@��
     k=k+1
     zgt(0)=rsgt(0)
       For i=1 to rsgt.Fields.Count-1
       if rsgt(i).Name= Trim(YKG) then
         'if rsgt(i).Name= "��" then
           'zgt(0)="�Ѥ~��"
           zgt(i)=rsgt(i)
           zgtk =rsgt(i)
     '   Select Case Trim(rsgt(0))
    ' Case "�Ѥ~"
     '      For n=5 to 17
      '         if rsgt(i)=A(n) then
      '            zgtk=B(n)
       '        end if                
        '     Next

     ' Case "�ѹ�" 
     '      BN5=Trim(B(3))
     '      YKGS=Trim(YKG)
      '   'zgtk= GGRNDSN(B(3), YKG)
      '   GGRNDSN BN5, YKGS

    '  End  Select      
        
         IF Trim(rsgt(0))= "�Ѥ~"  Then 
            For n=5 to 17
               if rsgt(i)=A(n) then
                  zgtk=B(n)
               end if                
             Next
          End If
           IF Trim(rsgt(0))= "�ѹ�"  Then 
               zgtk= GGRNDSN(B(3), YKG)
                BN5=Trim(B(3))
                YKGS=Trim(YKG)
              ' zgtk= GGRNDSN(BN3, YKGS)
             ' GGRNDSN BN3, YKGS

            End IF
       End If  
     Next
          
    ' rsgt.MoveNext	' ����U�@��
  ' Wend 
 
  A(0+k)=zgt(0)
  B(0+k)=zgtk
  AW(0+k)=zgt(0)& wk
  'AW(0+k)=zgt(0)&wk
   
 rsgt.MoveNext	' ����U�@��
 Wend
 
  'DOCTOR SexYKT, subtk
  'GGRNDSN BN3, YKGS

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %>
 
  <%  '�p��Ѥ~�ئ~��
 ' SUB GGRNDSN(BN5, YKGS)
 FUNCTION GGRNDSN(BN3, YKGS)
 
   ZGG= Array("�l","��","�G","�f","��","�w","��","��","��","��","��","��")
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
  ' A(105)="�Ѥ~��"
  ' B(105)=zgtk
 ' AW(105)=zgt(0)& wk

' END SUB
      
 END FUNCTION 

  
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  %>
   
  <%  ''�p��T�x�K�y�����ѶQ�c��
    'SUB TGRNDSN(TBNM, TDNUM)
   FUNCTION TGRNDSN(TBNM, TDNUM)
 
   ZGG= Array("�l","��","�G","�f","��","�w","��","��","��","��","��","��")
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
  ' A(0+k)=�T�x
  ' B(0+k)=zmtk
 ' AW(0+k)=zmt(0)& wk
  
' END SUB
      
 END FUNCTION 

   
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 %>   

<%
 ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From ��t�P��"
Set rsmt = GetExcelRecordset( "Excel09.xls",  SQL)
If rsmt Is Nothing Then
    Response.Write "GetExcelRecordset �I�s����!"
    Response.End
End If 
  ReDim zmt(16)
 ' zmt(0)="���k�]"
  'zmtk=" "
  wk=" "
  k=60


  ' Part II�G��X��ƪ��u���e�v
  rsmt.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
  While Not rsmt.EOF	' �P�_�O�_�L�F�̫�@��
     k=k+1
     zmt(0)=rsmt(0)
    For i=1 to rsmt.Fields.Count-1
         'if Trim(rsmt(i).Name)=Trim(TNUM) then
      if Trim(rsmt(i).Name)=Trim("M"&TNUM) then
         'if rsmt(i).Name= "��" then
           'zmt(0)="���k�]"
           zmt(i)=rsmt(i)
           zmtk =rsmt(i)
           
      Select Case Trim(zmt(0))
        Case "�T�x"
          TDNUM1=DNUM-1
          zmtk=TGRNDSN(B(61), TDNUM1)
           ' DOCTOR SexYKT, subtk
         Case "�K�y" 
           TDNUM2=-(DNUM-1)
           zmtk=TGRNDSN(B(62), TDNUM2)
            ' GGRNDSN BN5, YKGS
         Case "����"
          TDNUM3=DNUM-2
          zmtk=TGRNDSN(B(54), TDNUM3)
           ' DOCTOR SexYKT, subtk
         Case "�ѶQ" 
           TDNUM4=DNUM-2
           zmtk=TGRNDSN(B(55), TDNUM4)
            ' GGRNDSN BN5, YKGS
       End  Select      
      
      End If  
    Next
          
    ' rsmt.MoveNext	' ����U�@��
  ' Wend 
 
  A(0+k)=zmt(0)
  B(0+k)=zmtk
  AW(0+k)=zmt(0)& wk
  'AW(0+k)=zmt(0)&wk
   
 rsmt.MoveNext	' ����U�@��
 Wend
 
   ''DOCTOR SexYKT, subtk
   ''GGRNDSN BN3, YKGS
  ' TGRNDSN(TBNM, TDNUM)

%>


<%

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From ���Ū�"
Set rssk = GetExcelRecordset( "Excel09.xls",  SQL)
If rssk Is Nothing Then
    Response.Write "GetExcelRecordset �I�s����!"
    Response.End
End If 
  ReDim zsk(16)
 ' zsk(0)="���Ŧ~��"
  'zskk=" "
  wk=" "
  k=105

 ' Part II�G��X��ƪ��u���e�v
  rssk.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
  While Not rssk.EOF	' �P�_�O�_�L�F�̫�@��
   if  Trim(rssk(0))= Trim(YKG)  then
        ' zsk(0)=rssk(0)
       For i=1 to rssk.Fields.Count-1
        if Trim(rssk(i).Name)= Trim(YKK)  then
           'if rssk(i).Name= "��" then
           'zsk(i)="���Ůc��"
           zsk(i)=rssk(i)
           zskk =rssk(i)
         End If 
       Next
    ' rssk.Update
   ' End If  
   ' rssk.MoveNext	' ����U�@��
 ' Wend  
    
   'A(0+k)=zsk(0)
   A(0+k)="����"
   B(0+k)=zskk
  'AW(0+k)=zsk(0)&wk
   AW(0+k)="����"&wk
   AKI(0+k)="F"&k
  End if  
 rssk.MoveNext	' ����U�@��
 
 Wend
 
 A(106)="�Ѷ�"
   B(106)=B(12)
  'AW(0+k)=zsk(0)&wk
   AW(106)="�Ѷ�"&wk
   'A(12)="���Юc"
 A(107)="�Ѩ�"
   B(107)=B(10)
  'AW(0+k)=zsk(0)&wk
   AW(107)="�Ѩ�"&wk
   ''A(10)="�e�̮c"

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 %>

<%

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From �|�ƬP��"
Set rsk = GetExcelRecordset( "Excel09.xls",  SQL)
If rsk Is Nothing Then
    Response.Write "GetExcelRecordset �I�s����!"
    Response.End
End If 
  ReDim zfk(16)
 ' zfk(0)="�|�ƬP�c"
  zfkk=" "
  wk=" "
  k=140

 ' Part II�G��X��ƪ��u���e�v
  rsk.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
  While Not rsk.EOF	' �P�_�O�_�L�F�̫�@��
     k=k+1
   ' For j=1 to 11 
   ' IF rsk(0)= "�Q"  Then 
      'ReDim zfk(16)
     zfk(0)=rsk(0)
       For i=1 to rsk.Fields.Count-1
        if  rsk(i).Name= Trim(YKK)  then
         'if rsk(i).Name= z11k then
         'if rsk(i).Name= "��" then
           'zfk(0)="�|�ƬP�c"
           zfk(i)=rsk(i)
           zfkk =rsk(i)
         End If 
       Next
    ' rsk.Update
   ' End If  
   ' rsk.MoveNext	' ����U�@��
 ' Wend  
    
   A(0+k)=zfk(0)
   B(0+k)=zfkk
   AW(0+k)=zfk(0)&wk
   AKI(0+k)="F"&k
       

 
 rsk.MoveNext	' ����U�@��
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
 SQL = "Select * From �j����"
Set rsg = GetExcelRecordset( "Excel09.xls",  SQL)
If rsg Is Nothing Then
    Response.Write "GetExcelRecordset �I�s����!"
    Response.End
End If 
  ReDim zfl(24)
 ' zfl(0)="�j����c"
  zflk=" "
  wk=" "
  k=145

 ' Part II�G��X��ƪ��u���e�v
  rsg.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
  While Not rsg.EOF	' �P�_�O�_�L�F�̫�@��
     k=k+1
   ' For j=1 to 11 
   ' IF rsg(0)= "�Q"  Then 
      'ReDim zfl(16)
     zfl(0)=rsg(0)
       For i=1 to rsg.Fields.Count-1
        if  rsg(i).Name= Trim(SexYK+z2k)  then
         ' if  rsg(i).Name= Trim(YKK)  then
         'if rsg(i).Name= z11k then
         'if rsg(i).Name= "��" then
           'zfl(0)="�j����c"
           zfl(i)=rsg(i)
           zflk =rsg(i)
         End If 
       Next
    ' rsg.Update
   ' End If  
   ' rsg.MoveNext	' ����U�@��
 ' Wend  
    
  ' A(0+k)=zfl(0)
  ' B(0+k)=zflk
  ' AW(0+k)=zfl(0)&wk

  A(0+k)=zflk
   B(0+k)=zfl(0)
  AW(0+k)=zflk&wk
   AKI(0+k)="G"&k
       

 rsg.MoveNext	' ����U�@��
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
 SQL = "Select * From �p����D"
Set rsgs = GetExcelRecordset( "Excel09.xls",  SQL)
If rsgs Is Nothing Then
    Response.Write "GetExcelRecordset �I�s����!"
    Response.End
End If 
  ReDim zfms(24)
 ' zfms(0)="�p����c"
  zfmsk=" "
  wk=" "
  k=160

 ' Part II�G��X��ƪ��u���e�v
  rsgs.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
  While Not rsgs.EOF	' �P�_�O�_�L�F�̫�@��
     k=k+1
   ' For j=1 to 11 
   ' IF rsgs(0)= "�Q"  Then 
      'ReDim zfms(16)
     zfms(0)=rsgs(0)
       For i=1 to rsgs.Fields.Count-1
        if  rsgs(i).Name= Trim(Sex+YKG)  then
         ' if  rsgs(i).Name= Trim(YKK)  then
         'if rsgs(i).Name= z11k then
         'if rsgs(i).Name= "��" then
           'zfms(0)="�p����c"
           zfms(i)=rsgs(i)
           zfmsk =rsgs(i)
         End If 
       Next
    ' rsgs.Update
   ' End If  
   ' rsgs.MoveNext	' ����U�@��
 ' Wend  
    
   A(0+k)= zfms(0)
   B(0+k)=zfmsk
   AW(0+k)=zfms(0)&wk
    AKI(0+k)="SL"&k

  'A(0+k)=zfmsk
  'B(0+k)=zfms(0)
  'AW(0+k)=zfmsk&wk
  ' AKI(0+k)="G"&k
       

 
 rsgs.MoveNext	' ����U�@��
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
 
  R= Array(" ","[�l�c]","[���c]","[�G�c]","[�f�c]","[���c]","[�w�c]","[�Ȯc]","[���c]","[�Ӯc]","[���c]","[���c]","[��c]","[���c]")

  RE= Array(" "," "," "," "," "," "," "," "," "," "," "," "," "," ")

 KFE= Array(" "," "," "," "," "," "," "," "," "," "," "," "," "," ")

  KF= Array(" ","[�l�c]","[���c]","[�G�c]","[�f�c]","[���c]","[�w�c]","[�Ȯc]","[���c]","[�Ӯc]","[���c]","[���c]","[��c]","[���c]")

 KS= Array("�R�c","�S��","�ҩd","�l�k","�]��","�e��","�E��","����","�x�S","�Цv","�ּw","����","���L","�Ѿ�","�Ӷ�","�Z��","�ѦP","�G�s","�ѩ�","�ӳ�","�g�T","����","�Ѭ�","�ѱ�","�C��","�}�x","�ƸS","���v","�Ƭ�","�Ƨ�"," ")
  KI= Array("K0","K1","K2","K3","K4","K5","K6","K7","K8","K9","K10","K11","S0","S1","S2","S3","S4","S5","T0","T1","T2","T3","T4","T5","T6","T7"," "," "," "," "," ")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ''�����c�W 
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
 ''   rs.Open "�c�z��", cn, adOpenKeyset, adLockOptimistic
  
  SQL = "Select * From �c�z��"
 Set rsy4 = GetExcelRecordset( "Excel09.xls",  SQL)
  
 ' R = Array(" ", "[�l�c]", "[���c]", "[�G�c]", "[�f�c]", "[���c]", "[�w�c]", "[�Ȯc]", "[���c]", "[�Ӯc]", "[���c]", "[���c]", "[��c]", "[���c]")
  ReDim zt(13)
 zt(0) = "�c�z"
   ' Part II�G��X��ƪ��u���e�v
  rsy4.MoveFirst      ' �N�ثe��ƿ�����Ĥ@��
  While Not rsy4.EOF  ' �P�_�O�_�L�F�̫�@��
    'If rsy4(0) = "��" Then
    If rsy4(0) = Trim(YKK) Then
     'zt(0)="�c�z"
       For i = 1 To rsy4.Fields.Count - 1
         zt(i) = rsy4(i) + rsy4(i).Name
        'rsy4(i) =  rsy4(i).Name+rsy4(i)
        R(i) = "[" + zt(i) + "�c]"
       Next
    End If
    
     rsy4.MoveNext    ' ����U�@��
  Wend
  '' Text1.Text = R(5) + R(10)
  rsy4.Close
   '' cn.Close


 %>


<%

  ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From �y�~�ѬP��"
Set rsy5 = GetExcelRecordset( "Excel09.xls",  SQL)
If rsy5 Is Nothing Then
    Response.Write "GetExcelRecordset �I�s����!"
    Response.End
End If 
  ReDim zy5(16)
 ' zy5(0)="�y�~�R�c�ѬP"
  zy5k=" "
  k=175
  wk=" "

  ' Part II�G��X��ƪ��u���e�v
  rsy5.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
  While Not rsy5.EOF	' �P�_�O�_�L�F�̫�@��
         k=k+1
      zy5(0)=rsy5(0)
       For i=1 to rsy5.Fields.Count-1
         if rsy5(i).Name= Trim(LYG) then
          'if rsy5(i).Name= "�G" then
           'zy5(0)="�y�~�R�c�ѬP"
           zy5(i)=rsy5(i)
           zy5k =rsy5(i)
          
        End If 
       Next
    ' rsy5.Update
   ' End If  
   ' rsy5.MoveNext	' ����U�@��
 ' Wend  
  A(0+k)=zy5(0)
  B(0+k)=zy5k
  AW(0+k)=zy5(0)&wk
  AKI(0+k)="K"&k
   
  rsy5.MoveNext	' ����U�@��
 Wend

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' SQL = "Select * From [B3:N14]"
 SQL = "Select * From �y�~��g��"
Set rsy6 = GetExcelRecordset( "Excel09.xls",  SQL)
If rsy6 Is Nothing Then
    Response.Write "GetExcelRecordset �I�s����!"
    Response.End
End If 
  ReDim zy6(16)
 ' zy6(0)="�y�~��g�c"
  zy6k=" "
  k=212
  wk=" "

  ' Part II�G��X��ƪ��u���e�v
  rsy6.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
  While Not rsy6.EOF	' �P�_�O�_�L�F�̫�@��
       
    If TRIM(rsy6(0))=Trim(TNUM) Then 
         'zy6(0)=rsy6(0)
       For i=1 to rsy6.Fields.Count-1
         if rsy6(i).Name= Trim(LYG) then
          'if rsy6(i).Name= "�G" then
           'zy6(0)="���R�c"
           zy6(i)=rsy6(i)
          zy66k =rsy6(i)
         zy6k= LGRNDSN(zy66k, HNUM)
    
        End If 
       Next
    ' rsy6.Update
   End If  
   ' rsy6.MoveNext	' ����U�@��
 ' Wend  
  'A(0+k)=zy6(0)
  A(212)="����g"
  B(212)=zy6k
 ' AW(0+k)="����g"zy6(0)&wk
  AW(0+k)="����g"&wk
  AKI(0+k)="K"&k
   
  rsy6.MoveNext	' ����U�@��
 Wend
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
%>
  
  <%  '�p��y�~��g
 ' SUB LGRNDSN(MONG, HNUMR)
 FUNCTION LGRNDSN(MONG, HNUMR)
 
   ZGG= Array("�l","��","�G","�f","��","�w","��","��","��","��","��","��")
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
      Case "�l"
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

    

       Case "��"
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


       Case "�G"
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

             
      Case "�f"
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


      Case "��"
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


     Case "�w"
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

    
      Case "��"
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
       

       Case "��"
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

      Case "��"
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
           

       Case "��"
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


       Case "��"
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


      Case "��"
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
Row2 = "<TR>" & "<TD width=230>" & Trim(R(5)) & "</TD>"& "<TD width=230  RowSpan=2 ColSpan=2>"& Trim(R(13))& SECTM&TNUM&DNUM&HNUM&",�R�D:"&MYL&",���D:"&MYB& "</TD>"& "<TD width =230>" & Trim(R(10)) & "</TD>"& "</TR>"
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
 ' SQL = "Select * From ���ת�2"
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

  <CENTER><H2><FONT color=#FF0000>�Y�ݭn�R �L �� �� ��,�Х[�J<u>�K�O�|��</u>,  ����</FONT><HR></H2>
  <!--<CENTER><h2><a href="http://class.ruten.com.tw/user/index00.php?s=tang1206">�^����</a></h2>-->
  <CENTER><h2 align="left"><a href="http://61.222.248.199/HSU-fundb/Login-2r.asp" ><u>�[�J�|��</u></a></h2>


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



 
 <%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Web.Mail" %>

 <%
  
  Dim TNam, TNsx ,Nox 
  Dim TNbr, TNyr,Nob, Noy ,BNum, YNum ,RTNBY ,RBYNS,RBN, RYN


   ''No = Request("No")
  'No1 = Request("No1")
  'No2 = Request("No2")
  'No3 = Request("No3")
  'TNo1 = Request("Name.Text")
  'TNo2 = Request("Tel.Text")
  'TNo3 = Request("Addr.Text")
  TNam = Request("Name")
  TNsx = Request("Sex")
  TNbr=Request("BYNUM")
  TNyr=Request("LYNUM")
   Nox = TNsx
  Nob = Cint(TNbr)
  Noy = Cint(TNyr)


 ''<script Language="VB" runat="server">

 '' Sub Button_Click(sender As Object, e As EventArgs) 
      '' Msg.Text = ""
      '' Dim URL 
      'No1 = Request("NNo1")
   'No2 = Request("NNo2")
   'No3 = Request("NNo3")
 Dim R()= {" ","[���c]","[�[�c]","[�_�c]","[�S�c]","[���c]","[���c]","[�I�c]","[��c]","[���c]","[���c]"}
 Dim ZNAM()={"��","��","�[","�_","�S","��","��","�I","��","��"}
 Dim ZNUM()={"0","1","2","3","4","5","6","7","8","9"}
   Select Case TNsx
   Case "�k"
      if Nob >=2000 then
       BNum = 9-((Nob-200) mod 9)
      else
       BNum = 10-((Nob-100) mod 9)
  end if
      Case "�k"
    if Nob >=2000 then
      BNum = ((Nob-200-3) mod 9)
   else
      BNum = ((Nob-100-4) mod 9)
  end if
 
   End  Select 
  Select Case BNum
   Case 0
     BNum = 9
   Case 10
     BNum = 1
      'if BNum = 10 then
      '  BNum = 1
      'end if 
  End  Select 
      
   if Noy >=2000 then
    YNum = 9-((Noy-200) mod 9)
   else
    YNum =10-((Noy-100) mod 9)
  end if
    if YNum = 10 then
        YNum = 1
     end if     

   Response.Write (BNum & ";" & YNum &"<br>")

     '' RTNBY = QueryDataAndSendTo(BNum, YNum)
     '' RBYNS = Split( RTNBY ,";")
     '  RBN =RBYNS
     '  RYN =RBYNS

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' <script Language="VB" runat="server">
   '''' Function  QueryDataAndSendTo(XNam1,XNam2)
    
      ON ERROR RESUME NEXT
      Dim Conn As OleDbConnection, Cmd As OleDbCommand
      Dim Rdb,Rdy As OleDbDataReader, SQL As String, Body As String
      Dim mad11, mad22 As String
      Dim I,J, k ,BBK As integer
      Dim A(30) ,B(30)
       k=0
      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      
       ''Dim Database = "Data Source=" & Server.MapPath( "../ch15/Users.mdb" )
       'Dim Database = "Data Source=" & Server.MapPath( "UsersPwd.mdb" )
        'Dim Database = "Data Source=" & Server.MapPath( "/HSU-fundb/UsersPwd.mdb" )
      Dim Database = "Data Source=" & Server.MapPath( "/HSU-PK/UsersPwd.mdb" )
      Dim Dbpass = "Jet OLEDB:Database Password=kj6688"
      Conn = New OleDbConnection( Provider & ";" & DataBase & ";"&Dbpass )
      Conn.Open()    
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
  Dim  Zdb(9)
  Dim  Zdbk=" "
  k=1
       ' �ˬd Email �O�_�s�b
      ' SQL = "Select * From Users Where Email='" & Emad1.Text & "'"
      'SQL = "Select * From �K���� Where ����='" & Emad1.Text & "'"
       SQL = "Select * From �ͦ~�����"
      Cmd = New OleDbCommand( SQL, Conn )
      Rdb = Cmd.ExecuteReader()
    While Rdb.Read()
       If Rdb.Item(0)= BNum Then ' ��ܦ��@ Email �s�b
        'If Rdb.Item("����")=Emad1.Text Then ' ��ܦ��@ Email �s�b
         'If Rdb .Read() Then ' ��ܦ��@ Email �s�b
         ' Dim mad1 As string
       for i = 1 to Rdb.Fieldcount -1
        '' mad11 = Rdb.Item("��")
         '' if Rdb .GetName(i) = XNam2 then
          ''mad22 =Rdb.Item(i)
         ' Zdb(i)=Rdb.GetName(i)
         ' Zdbk =Rdb.Item(i)

          A(k+i) = "("& Rdb.GetName(i) & ")" &"<br>"
         B(k+i) = Rdb.Item(i)
 
         '' End If 

       '' Response.Write (A(k+i)& B(k+i))
      next 
    end if
   'A(0+k)=Zdb(i)
   'B(0+k)=Zdbk
  ''k=k+1
 End While   
     Rdb.close()
     '' Conn.Close()
  
  Dim  Zdy(9)
  Dim  Zdyk=" "
  k=10
       ' �ˬd Email �O�_�s�b
      ' SQL = "Select * From Users Where Email='" & Emad1.Text & "'"
      'SQL = "Select * From �K���� Where ����='" & Emad1.Text & "'"
       SQL = "Select * From �y�~���ժ�"
      Cmd = New OleDbCommand( SQL, Conn )
      Rdy = Cmd.ExecuteReader()
    While Rdy.Read()
       If Rdy.Item(0)= YNum Then ' ��ܦ��@ Email �s�b
        'If Rdy.Item("����")=Emad1.Text Then ' ��ܦ��@ Email �s�b
         'If Rdy .Read() Then ' ��ܦ��@ Email �s�b
         ' Dim mad1 As string
       for i = 1 to Rdy.Fieldcount -1
        '' mad11 = Rdy.Item("��")
         '' if Rdy .GetName(i) = XNam2 then
          ''mad22 =Rdy.Item(i)
         ' Zdy(i)=Rdy.GetName(i)
         '  Zdyk =Rdy.Item(i)

        A(k+i) = Rdy.GetName(i)
        B(k+i) = Rdy.Item(i)
 
         '' End If 
      '' Response.Write (A(k+i)& B(k+i))
     next 
    end if
   ''A(0+k)=Zdy(i)
  '' B(0+k)=Zdyk
  
   ''k=k+1
 
 End While   
     Rdy.close()
      Conn.Close()
   
  For j=0 to 20
      BBK=Trim(B(j))
   Select Case BBK
      Case "1"
        ''R(1)=Trim(R(1))+","& Trim(A(j))
        R(1)=Trim(R(1))+"<br>"& Trim(A(j))
      Case "2"
        R(2)=Trim(R(2))&"<br>"& Trim(A(j))  
      Case "3"
        R(3)=Trim(R(3))+"<br>"& Trim(A(j)) 
      Case "4"
        R(4)=Trim(R(4))+"<br>"& Trim(A(j))  
      Case "5"
        R(5)=Trim(R(5))+"<br>"& Trim(A(j))
      Case "6"
        R(6)=Trim(R(6))+"<br>"& Trim(A(j))  
      Case "7"
        R(7)=Trim(R(7))+"<br>"& Trim(A(j))
      Case "8"
        R(8)=Trim(R(8))+"<br>"& Trim(A(j))
      Case "9"
        R(9)=Trim(R(9))+"<br>"& Trim(A(j))
     Case Else
        R(10)=Trim(R(10))+"<br>"& Trim(A(j))        
   End  Select      
  Next      

  
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''      
 Dim Row1 ="<CENTER><TABLE Border=1 BGCOLOR=#FFFF00 >"&"<TR>"&"<TD width=120>���W"&"</TD>"&"<TD width=120>(�ͦ���)</TD>"&"<TD width =120>�y����</TD>"&"</TR>"
 Dim Row2 = "<TR>"& "<TD width=120>" & R(4) & "</TD>" & "<TD width=120>" & R(9) & "</TD>"&  "<TD width =120>" & R(2) & "</TD>"& "</TR>"
 Dim Row3 = "<TR>"& "<TD width=120>" & R(3) & "</TD>" & "<TD width=120>" & R(5) & "</TD>"&  "<TD width =120>" & R(7) & "</TD>"&"</TR>"
 Dim Row4 = "<TR>"& "<TD width=120>" & R(8) & "</TD>" & "<TD width=120>" & R(1) & "</TD>"&  "<TD width =120>" & R(6) & "</TD>"& "</TR>"
 Dim Row5 = "<TR style='writing-mode:bt-rl' width =150>" & "�L���ۧ@�v:<u>�K�O����</u>"& "</TR>"& "</TABLE>"
 Dim RowT=Row1+Row2+Row3+Row4+Row5
  Response.Write (RowT)
 
 %>
  <HTML> 
   
 <BODY bgcolor="#FFFFFF">
  <CENTER><H2> �� �g �K �� �L �P �� �� �� �p �U<HR></H2>
  <TABLE Border=1 Width=60%  BGCOLOR=#FFFF00>
   <!--<TR><TD width=30>���W</TD><TD width=30>����</TD><TD width =30>�ܤ�</TD><TD width =30>���W</TD><TD width =30>����</TD></TR>
   -->
  <%  
 ''Dim Row1 ,Row2 ,Row3 ,Row4 ,Row5 ,Row6 ,Row7 ,Row8 ,RowT
    ''    Response.Write (RowT)
  %>

</TABLE></CENTER>
 </BODY>
 
  <!-- <CENTER><H2><FONT color=#FF0000>�Y�ݭn<u>�� �L �� �� ��</u></FONT><HR></H2>
 <CENTER><H2><FONT color=#FF0000>    �� Email:  tech.t1206@gmail.com �Y�i�p��,  ����</FONT><HR></H2>
  -->
</Html>
 
 
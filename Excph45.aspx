 
 <%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Web.Mail" %>

 <%
  
     Dim TNam, TNsx, TNhs, InpLn, InpSFn, InpSn, InpFn
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
     InpSFn = Request("HSNUM")
     ''TNhs=Request("HSNUM")
     TNhs = Mid(InpSFn, 2, 1)
     InpSn = Mid(InpSFn, 2, 1)
     InpFn = Mid(InpSFn, 4, 1)
 Nob = CInt(TNbr)
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
     ''Dim ZNFS()={"��","�l","�[","�f","�S","��","��","��","��","��"}  
     ''Dim TNTFS() = {"0", "��", "�l", "��", "��", "��", "�G", "��", "�f", "�A", "��", "�S", "�x", "��", "��", "�B", "��", "�[", "��", "��", "��", "��", "��", "��", "��"}
     Dim TNTFS() = {"0", "��", "��", "��", "��", "��", "�G", "��", "�_", "�A", "��", "�S", "�x", "��", "��", "�B", "��", "�[", "��", "��", "�I", "��", "��", "��", "��"}
     Dim TNTFN() = {"0", "11", "12", "13", "81", "82", "83", "31", "32", "33", "41", "42", "43", "91", "92", "93", "21", "22", "23", "71", "72", "73", "61", "62", "63"}
     Dim TNTSX() = {"0", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��"}
     Select Case TNsx
         Case "�k"
             If Nob >= 1863 Then
                 BNum = (11 - ((Nob - 1863) Mod 9)) Mod 9
                 '' Else
                 ''  Response.Write("�Ȧ~�W�W���d��")
             End If
             If BNum = 5 Then
                 BNum = 2
             End If
         Case "�k"
             If Nob >= 1863 Then
                 BNum = (4 + ((Nob - 1863) Mod 9)) Mod 9
                 '' Else
                 ''   Response.Write("�Ȧ~�W�W���d��")
             End If
             If BNum = 5 Then
                 BNum = 8
             End If
     End Select
     If BNum = 0 Then
         BNum = 9
     End If
     
    
     If Noy >= 1863 Then
         YNum = (11 - ((Noy - 1863) Mod 9)) Mod 9
         ''Else
        '' Response.Write("�Ȧ~�W�W���d��")
     End If
     '' If YNum = 5 Then  ''�K�v��(�y)�~���~�ϥ�"
     ''Select Case TNsx
     ''    Case "�k"
     '' YNum = 2
     ''     Case "�k"
     ''  YNum = 8
     ''  End Select
     '''  End If
     If YNum = 0 Then
         YNum = 9
     End If

     Response.Write(BNum & ";" & YNum & "<br>")

     '' RTNBY = QueryDataAndSendTo(BNum, YNum)
     '' RBYNS = Split( RTNBY ,";")
     '  RBN =RBYNS
     '  RYN =RBYNS

     ''''''''''''''''''''''''''''''''''''''''''''''''''''''
     '''' <script Language="VB" runat="server">
     '''' Function  QueryDataAndSendTo(XNam1,XNam2)
    
     On Error Resume Next
     Dim Conn As OleDbConnection, Cmd As OleDbCommand
     Dim Rdb, Rdy, Rhs As OleDbDataReader, SQL As String, Body As String
     Dim mad11, mad22 As String
     Dim I, J, k, BBK As Integer
     Dim A(30), B(30)
     k = 0
     Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      
     ''Dim Database = "Data Source=" & Server.MapPath( "../ch15/Users.mdb" )
     'Dim Database = "Data Source=" & Server.MapPath( "UsersPwd.mdb" )
     'Dim Database = "Data Source=" & Server.MapPath( "/HSU-fundb/UsersPwd.mdb" )
     Dim Database = "Data Source=" & Server.MapPath("/HSU-PK/UsersPwd.mdb")
     Dim Dbpass = "Jet OLEDB:Database Password=kj6688"
     Conn = New OleDbConnection(Provider & ";" & Database & ";" & Dbpass)
     Conn.Open()
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
     Dim Zdb(9)
     Dim Zdbk = " "
     k = 1
     ' �ˬd Email �O�_�s�b
     ' SQL = "Select * From Users Where Email='" & Emad1.Text & "'"
     'SQL = "Select * From �K���� Where ����='" & Emad1.Text & "'"
     SQL = "Select * From �ͦ~�����"
     Cmd = New OleDbCommand(SQL, Conn)
     Rdb = Cmd.ExecuteReader()
     While Rdb.Read()
         If Rdb.Item(0) = BNum Then ' ��ܦ��@ Email �s�b
             'If Rdb.Item("����")=Emad1.Text Then ' ��ܦ��@ Email �s�b
             'If Rdb .Read() Then ' ��ܦ��@ Email �s�b
             ' Dim mad1 As string
             For I = 1 To Rdb.FieldCount - 1
                 '' mad11 = Rdb.Item("��")
                 '' if Rdb .GetName(i) = XNam2 then
                 ''mad22 =Rdb.Item(i)
                 ' Zdb(i)=Rdb.GetName(i)
                 ' Zdbk =Rdb.Item(i)

                 A(k + I) = "(" & Rdb.GetName(I) & ")"
                 B(k + I) = Rdb.Item(I)
 
                 '' End If 

                 '' Response.Write (A(k+i)& B(k+i))
             Next
         End If
         'A(0+k)=Zdb(i)
         'B(0+k)=Zdbk
         ''k=k+1
     End While
     Rdb.Close()
     '' Conn.Close()
  
     Dim Zdy(9)
     Dim Zdyk = " "
     k = 10
     ' �ˬd Email �O�_�s�b
     ' SQL = "Select * From Users Where Email='" & Emad1.Text & "'"
     'SQL = "Select * From �K���� Where ����='" & Emad1.Text & "'"
     SQL = "Select * From �y�~���ժ�"
     Cmd = New OleDbCommand(SQL, Conn)
     Rdy = Cmd.ExecuteReader()
     While Rdy.Read()
         If Rdy.Item(0) = YNum Then ' ��ܦ��@ Email �s�b
             'If Rdy.Item("����")=Emad1.Text Then ' ��ܦ��@ Email �s�b
             'If Rdy .Read() Then ' ��ܦ��@ Email �s�b
             ' Dim mad1 As string
             For I = 1 To Rdy.FieldCount - 1
                 '' mad11 = Rdy.Item("��")
                 '' if Rdy .GetName(i) = XNam2 then
                 ''mad22 =Rdy.Item(i)
                 ' Zdy(i)=Rdy.GetName(i)
                 '  Zdyk =Rdy.Item(i)

                 A(k + I) = Rdy.GetName(I)
                 B(k + I) = Rdy.Item(I)
                 If Rdy.Item(I) = 5 Then
                     A(k + I) = "�y�~�P:" & A(k + I)
                 End If
                 '' End If 
                 '' Response.Write (A(k+i)& B(k+i))
             Next
         End If
         ''A(0+k)=Zdy(i)
         '' B(0+k)=Zdyk
  
         ''k=k+1
 
     End While
     Rdy.Close()
     '' Conn.Close()
   
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     Dim Zhs(9)
     Dim Zhsk = " "
     k = 20
     ' �ˬd Email �O�_�s�b
     ' SQL = "Select * From Users Where Email='" & Emad1.Text & "'"
     'SQL = "Select * From �K���� Where ����='" & Emad1.Text & "'"
     SQL = "Select * From �v��������"
     '' SQL = "Select * From �v�����ժ�"
     Cmd = New OleDbCommand(SQL, Conn)
     Rhs = Cmd.ExecuteReader()
     While Rhs.Read()
         If Rhs.Item(0) = TNhs Then ' ��ܦ��@ Email �s�b
             'If Rhs.Item("����")=Emad1.Text Then ' ��ܦ��@ Email �s�b
             'If Rhs .Read() Then ' ��ܦ��@ Email �s�b
             ' Dim mad1 As string
             For I = 1 To Rhs.FieldCount - 1
                 '' mad11 = Rhs.Item("��")
                 '' if Rhs .GetName(i) = XNam2 then
                 ''mad22 =Rhs.Item(i)
                 B(k + I) = Rhs.GetName(I)
                 A(k + I) = Rhs.Item(I)
                 If Rhs.GetName(I) = 5 Then
                     A(k + I) = "���s�P:" & A(k + I)
                 End If
                 '' A(k + I) = Rhs.GetName(I) ''B()�s����,A()�s����
                 '' B(k + I) = Rhs.Item(I)   '' �v��������P���ժ��� 
                 '' End If 
                 '' Response.Write (A(k+i)& B(k+i))
             Next
         End If
         ''A(0+k)=Zhs(i)
         '' B(0+k)=Zhsk
  
         ''k=k+1
 
     End While
     Rhs.Close()
     Conn.Close()

     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
     For J = 0 To 30
         BBK = Trim(B(J))
         Select Case BBK
             Case "1"
                 ''R(1)=Trim(R(1))+","& Trim(A(j))
                 R(1) = Trim(R(1)) + "<br>" & Trim(A(J))
             Case "2"
                 R(2) = Trim(R(2)) & "<br>" & Trim(A(J))
             Case "3"
                 R(3) = Trim(R(3)) + "<br>" & Trim(A(J))
             Case "4"
                 R(4) = Trim(R(4)) + "<br>" & Trim(A(J))
             Case "5"
                 R(5) = Trim(R(5)) + "<br>" & Trim(A(J))
             Case "6"
                 R(6) = Trim(R(6)) + "<br>" & Trim(A(J))
             Case "7"
                 R(7) = Trim(R(7)) + "<br>" & Trim(A(J))
             Case "8"
                 R(8) = Trim(R(8)) + "<br>" & Trim(A(J))
             Case "9"
                 R(9) = Trim(R(9)) + "<br>" & Trim(A(J))
             Case Else
                 R(10) = Trim(R(10)) + "<br>" & Trim(A(J))
         End Select
     Next

  
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''      
     Dim Row1 = "<CENTER><TABLE Border=1 BGCOLOR=#FFFF00 >" & "<TR>" & "<TD width=120>���W" & "</TD>" & "<TD width=120>(�ͦ���)</TD>" & "<TD width =120>�y����</TD>" & "</TR>"
     Dim Row2 = "<TR>" & "<TD width=120>" & R(4) & "</TD>" & "<TD width=120>" & R(9) & "</TD>" & "<TD width =120>" & R(2) & "</TD>" & "</TR>"
     Dim Row3 = "<TR>" & "<TD width=120>" & R(3) & "</TD>" & "<TD width=120>" & R(5) & "</TD>" & "<TD width =120>" & R(7) & "</TD>" & "</TR>"
     Dim Row4 = "<TR>" & "<TD width=120>" & R(8) & "</TD>" & "<TD width=120>" & R(1) & "</TD>" & "<TD width =120>" & R(6) & "</TD>" & "</TR>"
     Dim Row5 = "<TR style='writing-mode:bt-rl' width =150>" & "�L���ۧ@�v:<u>�K�O����</u>" & "</TR>" & "</TABLE>"
     Dim RowT = Row1 + Row2 + Row3 + Row4 + Row5
     Response.Write(RowT)
 
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
 
 
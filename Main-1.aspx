<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
 <%@ Page Language="VB" Debug="true" %>

<%
   ''Dim  Driver , DBpath,Param  ''Dim conn, rs
   ''Dim conn As OleDbConnection, Cmd As OleDbCommand

 ' Driver = "Driver={Microsoft Excel Driver (*.xls)};"
 ' DBPath = "DBQ=" & Server.MapPath("TEST.XLS")
 ' Param =Driver & "ReadOnly=0;" & DBPath
   ''Param =Driver& DBPath
    ''Dim SQLL As String 
 '  SQLL = "Select * From ��q��"
  ' conn = Server.CreateObject("ADODB.Connection")
  '  conn.Open Param
   '  Set rs = Server.CreateObject("ADODB.Recordset")
   ' rs.Open SQLL, conn, 2, 2
  %>   
<HTML>
<BODY Background="../B01.jpg">

<Form runat="server">
<H2>KJ �q�l�ʪ��� -- ����ʪ���<HR></H2>
<BLOCKQUOTE>
<asp:RadioButtonList runat="server" id="Category" />
<asp:Button runat="server" Text="�i�J�ʪ���" OnClick="Enter_Click" />
</BLOCKQUOTE>
<HR><asp:Label runat="server" id="Msg" ForeColor=Red />
<Table Width=250><Tr>
<Td Align=Center><A HREF="List.aspx">�d���ʪ��U</A></Td>
<Td Align=Center><A HREF="Clear.aspx">�h�^�Ҧ��ӫ~</A></Td>
</Tr></Table>
</Form>

</BODY>
</HTML>

<script language="VB" runat=server>

   Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"

   '' Dim Provider = "Driver={Microsoft Excel Driver (*.xls)};"
   ''Dim Database = "Data Source=" & Server.MapPath("eShop.mdb" )
    Dim Database = "Data Source=" & Server.MapPath("TEST.XLS" )
 
   Sub Page_Load(sender As Object, e As EventArgs)
      If Not IsPostBack Then
         Session("IsCookieOpen") = "Set in Page_Load"
      
         Dim Conn As OleDbConnection, Cmd As OleDbCommand
         Dim Rd As OleDbDataReader, SQL As String ,Row As String
         Conn = New OleDbConnection( Provider & ";" & DataBase )
         Conn.Open()
         SQL = "Select * From ��q��"
       ''SQL = "Select * From ���O"
         Cmd = New OleDbCommand( SQL, Conn )
         Rd = Cmd.ExecuteReader()
         Dim I As Integer
         While Rd.Read()
            Category.Items.Add( Rd.Item("���O�W��") )
            Category.Items(I).Value = Rd.Item("���O�s��")
            I += 1
           Row = Row+"<TD>" & Rd(I) & "</TD>"
          '' Response.Write (Row)
           End While 
         
         Conn.Close()
      Response.Write ( Row)
     
      End If
   End Sub

   Sub Enter_Click(sender As Object, e As EventArgs)
      Msg.Text = ""
     ' Row = "<TD>" & rs(6) & "</TD>"
    ' Response.Write Row 
  
      
   End Sub

</script>
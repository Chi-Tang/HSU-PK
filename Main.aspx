<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>

<HTML>
<BODY Background="../B01.jpg">

<Form runat="server">
<H2>KJ �q�l�ʪ��� -- ����ʪ���<HR></H2>
<BLOCKQUOTE>
<asp:RadioButtonList runat="server" id="Category" /><p>
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
   Dim Database = "Data Source=" & Server.MapPath( "eShop.mdb" )

   Sub Page_Load(sender As Object, e As EventArgs)
      If Not IsPostBack Then
         Session("IsCookieOpen") = "Set in Page_Load"
      
         Dim Conn As OleDbConnection, Cmd As OleDbCommand
         Dim Rd As OleDbDataReader, SQL As String
         Conn = New OleDbConnection( Provider & ";" & DataBase )
         Conn.Open()
         SQL = "Select * From ���O"
         Cmd = New OleDbCommand( SQL, Conn )
         Rd = Cmd.ExecuteReader()
         Dim I As Integer
         While Rd.Read()
            Category.Items.Add( Rd.Item("���O�W��") )
            Category.Items(I).Value = Rd.Item("���O�s��")
            I += 1
         End While 
         Conn.Close()
      End If
   End Sub

   Sub Enter_Click(sender As Object, e As EventArgs)
      Msg.Text = ""
      
      If Session("IsCookieOpen") <> "Set in Page_Load" Then
         Msg.Text = "���i�J�ʪ�, �Х��}���s������ Cookie, " & _
                    "�M�������s����, �A���s�Ұ��s����!"
         Exit Sub
      End If
      
      Dim Sel = Category.SelectedItem
      If Not Sel Is Nothing Then
         Response.Redirect( "Buy.aspx?���O�W��=" & Sel.Text & _
                            "&���O�s��=" & Sel.Value )
      End If
   End Sub

</script>
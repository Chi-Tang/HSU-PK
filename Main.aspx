<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>

<HTML>
<BODY Background="../B01.jpg">

<Form runat="server">
<H2>KJ 電子購物街 -- 選擇購物區<HR></H2>
<BLOCKQUOTE>
<asp:RadioButtonList runat="server" id="Category" /><p>
<asp:Button runat="server" Text="進入購物區" OnClick="Enter_Click" />
</BLOCKQUOTE>
<HR><asp:Label runat="server" id="Msg" ForeColor=Red />
<Table Width=250><Tr>
<Td Align=Center><A HREF="List.aspx">查看購物袋</A></Td>
<Td Align=Center><A HREF="Clear.aspx">退回所有商品</A></Td>
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
         SQL = "Select * From 類別"
         Cmd = New OleDbCommand( SQL, Conn )
         Rd = Cmd.ExecuteReader()
         Dim I As Integer
         While Rd.Read()
            Category.Items.Add( Rd.Item("類別名稱") )
            Category.Items(I).Value = Rd.Item("類別編號")
            I += 1
         End While 
         Conn.Close()
      End If
   End Sub

   Sub Enter_Click(sender As Object, e As EventArgs)
      Msg.Text = ""
      
      If Session("IsCookieOpen") <> "Set in Page_Load" Then
         Msg.Text = "欲進入購物, 請先開啟瀏覽器的 Cookie, " & _
                    "然後關閉瀏覽器, 再重新啟動瀏覽器!"
         Exit Sub
      End If
      
      Dim Sel = Category.SelectedItem
      If Not Sel Is Nothing Then
         Response.Redirect( "Buy.aspx?類別名稱=" & Sel.Text & _
                            "&類別編號=" & Sel.Value )
      End If
   End Sub

</script>
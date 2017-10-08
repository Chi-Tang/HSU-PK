<!-- #include virtual="/kjasp/func/DB.fun" -->
<HTML>
<BODY BgColor=White>
<H2>開啟 SQL Server 的 Northwind 資料庫的 Products 資料表<HR></H2>
<%

If Request("Send") <> Empty Then
   Set conn = GetSQLServerConnection( Request("DataSource"), Request("UserID"), Request("Password"), "Northwind" )
   Set rs = GetSQLServerStaticRecordset( conn, "Products" )
%>
<TABLE BORDER=1>
<TR BGCOLOR=#00FFFF>
<%
  ' Part I：輸出「抬頭名稱」
  For i=0 to rs.Fields.Count-1
     Response.Write "<TD>" & rs(i).Name & "</TD>"
  Next
%>
</TR>
<%
  ' Part II：輸出資料表的「內容」
  rs.MoveFirst		' 將目前資料錄移到第一筆
  While Not rs.EOF	' 判斷是否過了最後一筆
     Row = "<TR>"
     For i=0 to rs.Fields.Count-1
        Row = Row & "<TD>" & rs(i) & "</TD>"
     Next
     Response.Write Row & "</TR>"
     rs.MoveNext	' 移到下一筆
  Wend
Else
%>

連接 SQL Server 之前，請輸入相關資料：
<FORM Action=UseSQL.asp Method=GET>
<TABLE>
<TR><TD>機器名稱：</TD><TD><INPUT Type=Text Name="DataSource"></TD></TR>
<TR><TD>使用者ID：</TD><TD><INPUT Type=Text Name="UserID"></TD></TR>
<TR><TD>密碼：</TD><TD><INPUT Type=Password Name="Password"></TD></TR>
</TABLE>
<INPUT Type=Submit Value=" 連 接 " Name="Send">
</FORM>
<%End If%>
</BODY>
</HTML>
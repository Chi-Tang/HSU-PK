<%
' 建立 Connection 物件
Set conn = Server.CreateObject("ADODB.Connection")
Provider = "Provider=Microsoft.Jet.OLEDB.4.0;"
DBPath = "Data Source=" & Server.MapPath( "Sample.mdb" )

' 連結資料庫 
conn.Open Provider & DBPath

Set rs = Server.CreateObject("ADODB.Recordset")
' 開啟資料來源, 參數二為 Connection物件
rs.Open "TEMP", conn, Cursor, 2
%>

<HTML>
<BODY bgcolor="#FFFFFF">
<CENTER><H2>自己建立 Connection 物件，以開啟資料庫<HR></H2>
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
%>
</TABLE></CENTER>
</BODY>
</HTML>
<%
' 建立 Connection 物件
Set conn = Server.CreateObject("ADODB.Connection")
Driver = "Driver={Microsoft Excel Driver (*.xls)};"
DBPath = "DBQ=" & Server.MapPath( "Sample.xls" )

' 呼叫 Open 方法連結資料庫
conn.Open Driver & DBPath

Set rs = Server.CreateObject("ADODB.Recordset")
' 開啟資料來源, 參數二為 Connection物件
rs.Open "Select * From [成績單$]", conn, 2, 2
%>

<HTML>
<BODY bgcolor="#FFFFFF">
<CENTER><H2>自己建立 Connection 物件，以開啟資料庫 -- Excel 版本<HR></H2>
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
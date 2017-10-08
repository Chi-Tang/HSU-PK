<!-- #include virtual="/kjasp/func/DB.fun" -->
<%
SQL = "Select * From [B3:G6]"
Set rs = GetExcelRecordset( "Excel02.xls",  SQL)
If rs Is Nothing Then
    Response.Write "GetExcelRecordset 呼叫失敗!"
    Response.End
End If 
%>
<HTML>
<BODY bgcolor="#FFFFFF">
<CENTER><H2>以 Excel 工作表的部分範圍為資料表<HR></H2>
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
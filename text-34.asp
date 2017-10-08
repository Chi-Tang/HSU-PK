<!-- #include virtual="/Hsu-pk/DB.fun" -->
<%
 'Dim  Driver , DBpath,Param
 ' Dim conn, rs
  Driver = "Driver={Microsoft Excel Driver (*.xls)};"
   DBPath = "DBQ=" & Server.MapPath("Excel02.xls")
   Param =Driver & "ReadOnly=0;" & DBPath
   SQL = "Select * From [B3:G6]"

    '  Set GetExcelConnection = GetConnection(Driver & "ReadOnly=0;" & DBPath)
    ' Dim conn
    ' On Error Resume Next  ''If Err.Number <> 0 Then Exit Function
    ' Set GetConnection = Nothing  
   Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open Param
   Set rs = Server.CreateObject("ADODB.Recordset")
   rs.Open SQL, conn, 2, 2


   
'' /SQL = "Select * From [B3:G6]"
'' /Set rs = GetExcelRecordset( "Excel02.xls",  SQL)
''If rs Is Nothing Then
   ' Response.Write "GetExcelRecordset 呼叫失敗!"
   ' Response.End
''  End If 
'Set conn = GetMdbConnection("Test1.mdb")
 ' Set cmd = Server.CreateObject( "ADODB.Command" )
 ' Set cmd.ActiveConnection = rs.ActiveConnection
  'SQLS ="Select * into ASP FROM ASPK"
  ' SQLS ="Select * into "&Lesson& " FROM " &Lesson&"K" 　
  ''SQLD ="Delete From "&Lesson
  ''cmd.CommandText = SQLD
   ''  cmd.Execute


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
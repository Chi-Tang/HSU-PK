<!-- #include virtual="/Hsu-pk/DB.fun" -->
<%
  N = NOW
  Nyy = DatePart("yyyy",N)
  Nmm = DatePart("m",N)
  Ndd = DatePart("d",N)
  Nymd = Nyy&(Nmm-1)&Ndd
 'Dim  Driver , DBpath,Param
 ' Dim conn, rs
  Driver = "Driver={Microsoft Excel Driver (*.xls)};"
  ''DBPath = "DBQ=" & Server.MapPath("Exce201308.xls")
  '' DBPath = "DBQ=" & Server.MapPath("TEST"&Nymd&".xls")
   DBPath = "DBQ=" & Server.MapPath("TEST201308.xls")
   Param =Driver & "ReadOnly=0;" & DBPath
   SQL = "Select * From [A1:I30]"
   ''SQL = "Select * From 日益表"
    '  Set GetExcelConnection = GetConnection(Driver & "ReadOnly=0;" & DBPath)
    ' Dim conn
    ' On Error Resume Next  ''If Err.Number <> 0 Then Exit Function
    ' Set GetConnection = Nothing  
   Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open Param
   Set rs = Server.CreateObject("ADODB.Recordset")
   rs.Open SQL, conn, 2, 2

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
  
   Response.Write Nymd
%>
</TABLE></CENTER>
</BODY>
</HTML>
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
   ' Response.Write "GetExcelRecordset �I�s����!"
   ' Response.End
''  End If 
'Set conn = GetMdbConnection("Test1.mdb")
 ' Set cmd = Server.CreateObject( "ADODB.Command" )
 ' Set cmd.ActiveConnection = rs.ActiveConnection
  'SQLS ="Select * into ASP FROM ASPK"
  ' SQLS ="Select * into "&Lesson& " FROM " &Lesson&"K" �@
  ''SQLD ="Delete From "&Lesson
  ''cmd.CommandText = SQLD
   ''  cmd.Execute


%>
<HTML>
<BODY bgcolor="#FFFFFF">
<CENTER><H2>�H Excel �u�@�������d�򬰸�ƪ�<HR></H2>
<TABLE BORDER=1>
<TR BGCOLOR=#00FFFF>
<%
  ' Part I�G��X�u���Y�W�١v
  For i=0 to rs.Fields.Count-1
     Response.Write "<TD>" & rs(i).Name & "</TD>"
  Next
%>
</TR>
<%
  ' Part II�G��X��ƪ��u���e�v
  rs.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
  While Not rs.EOF	' �P�_�O�_�L�F�̫�@��
     Row = "<TR>"
     For i=0 to rs.Fields.Count-1
        Row = Row & "<TD>" & rs(i) & "</TD>"
     Next
     Response.Write Row & "</TR>"
     rs.MoveNext	' ����U�@��
  Wend
%>
</TABLE></CENTER>
</BODY>
</HTML>
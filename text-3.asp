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
   ''SQL = "Select * From ��q��"
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
  
   Response.Write Nymd
%>
</TABLE></CENTER>
</BODY>
</HTML>
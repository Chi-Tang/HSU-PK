<%
' �إ� Connection ����
Set conn = Server.CreateObject("ADODB.Connection")
Driver = "Driver={Microsoft Excel Driver (*.xls)};"
DBPath = "DBQ=" & Server.MapPath( "Sample.xls" )

' �I�s Open ��k�s����Ʈw
conn.Open Driver & DBPath

Set rs = Server.CreateObject("ADODB.Recordset")
' �}�Ҹ�ƨӷ�, �ѼƤG�� Connection����
rs.Open "Select * From [���Z��$]", conn, 2, 2
%>

<HTML>
<BODY bgcolor="#FFFFFF">
<CENTER><H2>�ۤv�إ� Connection ����A�H�}�Ҹ�Ʈw -- Excel ����<HR></H2>
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
<!-- #include virtual="/kjasp/func/DB.fun" -->
<%
SQL = "Select * From [B3:G6]"
Set rs = GetExcelRecordset( "Excel02.xls",  SQL)
If rs Is Nothing Then
    Response.Write "GetExcelRecordset �I�s����!"
    Response.End
End If 
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
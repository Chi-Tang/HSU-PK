<!-- #include virtual="/kjasp/func/DB.fun" -->
<HTML>
<BODY BgColor=White>
<H2>�}�� SQL Server �� Northwind ��Ʈw�� Products ��ƪ�<HR></H2>
<%

If Request("Send") <> Empty Then
   Set conn = GetSQLServerConnection( Request("DataSource"), Request("UserID"), Request("Password"), "Northwind" )
   Set rs = GetSQLServerStaticRecordset( conn, "Products" )
%>
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
Else
%>

�s�� SQL Server ���e�A�п�J������ơG
<FORM Action=UseSQL.asp Method=GET>
<TABLE>
<TR><TD>�����W�١G</TD><TD><INPUT Type=Text Name="DataSource"></TD></TR>
<TR><TD>�ϥΪ�ID�G</TD><TD><INPUT Type=Text Name="UserID"></TD></TR>
<TR><TD>�K�X�G</TD><TD><INPUT Type=Password Name="Password"></TD></TR>
</TABLE>
<INPUT Type=Submit Value=" �s �� " Name="Send">
</FORM>
<%End If%>
</BODY>
</HTML>
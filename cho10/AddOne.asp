<!-- #include virtual="/kjasp/func/DB.fun" -->
<%
' ���إ� Connection ����
Set conn = GetMdbConnection( "Sample.mdb" )

' �إ� Command ����, �ó]�w������ Connection ����
Set cmd = Server.CreateObject( "ADODB.Command" )
Set cmd.ActiveConnection = conn

SQL = "Update �Ѳ��污�� Set ����q = ����q + 1"
cmd.CommandText = SQL
cmd.Execute
%>
�N�u�Ѳ��污��v�Ҧ���ƿ����u����q�v�[�@�C
<!-- #include virtual="/kjasp/func/DB.fun" -->
<%
' 先建立 Connection 物件
Set conn = GetMdbConnection( "Sample.mdb" )

' 建立 Command 物件, 並設定對應之 Connection 物件
Set cmd = Server.CreateObject( "ADODB.Command" )
Set cmd.ActiveConnection = conn

SQL = "Update 股票行情表 Set 成交量 = 成交量 + 1"
cmd.CommandText = SQL
cmd.Execute
%>
將「股票行情表」所有資料錄的「成交量」加一。
<!-- #include virtual="/kjasp/func/DB.fun" -->
<%
Msg = "尚未輸入指令"
DB = Request("DB")
If DB = Empty Then DB = "Sample.mdb"
SQL = Request("SQL")

On Error Resume Next
If SQL <> Empty Then
   Set conn = GetMdbConnection( DB )
   Set cmd = Server.CreateObject( "ADODB.Command" )

   Set cmd.ActiveConnection = conn
   cmd.CommandText = SQL
   cmd.Execute

   If Err.Number = 0 Then 
      Msg = "OK"
   Else
      Msg = Err.Description
   End If
End If
%>

<HTML>
<BODY bgcolor="#FFFFFF">
<H2>TestCmd.asp -- Action Query 測試程式<HR></H2>
<FORM Action=TestCmd.asp Method=POST>
資料庫：<INPUT Type=Text Name=DB Value="<%=DB%>"><BR>
指　令：<INPUT Type=Text Name=SQL Size=50 Value="<%=SQL%>"><P>
<INPUT Type=Submit Value=" 執 行 ">
</FORM>
<HR>指令執行情況: <FONT Color=Red><%=Msg%>!</FONT>
</BODY>
</HTML>
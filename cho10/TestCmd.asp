<!-- #include virtual="/kjasp/func/DB.fun" -->
<%
Msg = "�|����J���O"
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
<H2>TestCmd.asp -- Action Query ���յ{��<HR></H2>
<FORM Action=TestCmd.asp Method=POST>
��Ʈw�G<INPUT Type=Text Name=DB Value="<%=DB%>"><BR>
���@�O�G<INPUT Type=Text Name=SQL Size=50 Value="<%=SQL%>"><P>
<INPUT Type=Submit Value=" �� �� ">
</FORM>
<HR>���O���污�p: <FONT Color=Red><%=Msg%>!</FONT>
</BODY>
</HTML>
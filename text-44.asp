<!-- #include virtual="/Hsu-pk/DB.fun" -->
<%


 ''Dim  Driver , DBpath,Param
 ' Dim conn, rs
  Driver = "Driver={Microsoft Excel Driver (*.xls)};"
 ' DBPath = "DBQ=" & Server.MapPath("TEST.XLS")
 ' Param =Driver & "ReadOnly=0;" & DBPath
   ''Param =Driver& DBPath
  ' SQL = "Select * From ��q��"
  '  Set conn = Server.CreateObject("ADODB.Connection")
  '  conn.Open Param
  '   Set rs = Server.CreateObject("ADODB.Recordset")
   ' rs.Open SQL, conn, 2, 2
  ''''''''''''''''''''''''''''''''''''''''''''''''''   
  ' Set cmd = Server.CreateObject( "ADODB.Command" )
  ' Set cmd.ActiveConnection = conn
    ''SQLD ="Delete From TEST.xls"
   '  SQLD ="Delete From ��q��"
   ' cmd.CommandText = SQLD
    '' cmd.Execute
'''''''''''''''''''''''''''''''''''''''''''''''''''''
 'Set conn = GetMdbConnection("Test1.mdb")
 ' Set cmd = Server.CreateObject( "ADODB.Command" )
 ' Set cmd.ActiveConnection = rs.ActiveConnection
  'SQLD ="Delete From Excel02.xls"
  ' cmd.CommandText = SQLD
   '  cmd.Execute

 Set conn = GetMdbConnection("TestAC-1.mdb")
  Set cmd = Server.CreateObject( "ADODB.Command" )
    Set rs = Server.CreateObject("ADODB.Recordset")
     Set cmd.ActiveConnection = rs.ActiveConnection
  'SQLS ="Select * into ASP FROM ASPK"
  ' SQLS ="Select * into "&Lesson& " FROM " &Lesson&"K" �@
  SQLD ="Delete From MATHA"
  cmd.CommandText = SQLD
     cmd.Execute


%>
<HTML>
<BODY bgcolor="#FFFFFF">
<CENTER><H2>�H Excel �u�@�������d�򬰸�ƪ�<HR></H2>
 <Script language="vb" >

Dim rngScriptAnchorRange As Range
Dim objNewScript As Script

Set rngScriptAnchorRange = ActiveWorkbook. _
    Worksheets(1).Range("B5")
Set objNewScript = ActiveWorkbook. _
    Worksheets(1).Scripts.Add(rngScriptAnchorRange, _
      msoScriptLocationInBody, _
      msoScriptLanguageVisualBasic, _
      "MyNewScript", , _
      "MsgBox (""Added Script object MyNewScript"")")
			
//�M�Ψ� SharedWorkspaceFiles ����C

//�U�C�d�ҷ|�s�W�ɮר�@�Τu�@�Ϫ��ɮ׶��X�C
    Dim swsfile As Office.SharedWorkspaceFile
    Set swsfile = ActiveWorkbook.SharedWorkspace.Files.Add( _
        "C:\MyWorkbook.xls", _
        , True, True)
    MsgBox "New file URL: " & swsfile.URL, _
        vbInformation + vbOKOnly, _
        "New File in Shared Workspace Files"
    Set swsfile = Nothing
</Script>


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
    ''' rs.Delete
     rs.MoveNext	' ����U�@��
  Wend
%>

</TABLE></CENTER>
</BODY>
</HTML>
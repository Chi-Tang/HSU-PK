<!-- #include virtual="/Hsu-pk/DB.fun" -->
<%


 ''Dim  Driver , DBpath,Param
 ' Dim conn, rs
  Driver = "Driver={Microsoft Excel Driver (*.xls)};"
 ' DBPath = "DBQ=" & Server.MapPath("TEST.XLS")
 ' Param =Driver & "ReadOnly=0;" & DBPath
   ''Param =Driver& DBPath
  ' SQL = "Select * From 日益表"
  '  Set conn = Server.CreateObject("ADODB.Connection")
  '  conn.Open Param
  '   Set rs = Server.CreateObject("ADODB.Recordset")
   ' rs.Open SQL, conn, 2, 2
  ''''''''''''''''''''''''''''''''''''''''''''''''''   
  ' Set cmd = Server.CreateObject( "ADODB.Command" )
  ' Set cmd.ActiveConnection = conn
    ''SQLD ="Delete From TEST.xls"
   '  SQLD ="Delete From 日益表"
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
  ' SQLS ="Select * into "&Lesson& " FROM " &Lesson&"K" 　
  SQLD ="Delete From MATHA"
  cmd.CommandText = SQLD
     cmd.Execute


%>
<HTML>
<BODY bgcolor="#FFFFFF">
<CENTER><H2>以 Excel 工作表的部分範圍為資料表<HR></H2>
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
			
//套用到 SharedWorkspaceFiles 物件。

//下列範例會新增檔案到共用工作區的檔案集合。
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
  ' Part I：輸出「抬頭名稱」
  For i=0 to rs.Fields.Count-1
     Response.Write "<TD>" & rs(i).Name & "</TD>"
  Next
%>
</TR>
<%
  ' Part II：輸出資料表的「內容」
  rs.MoveFirst		' 將目前資料錄移到第一筆
  While Not rs.EOF	' 判斷是否過了最後一筆
     Row = "<TR>"
     For i=0 to rs.Fields.Count-1
        Row = Row & "<TD>" & rs(i) & "</TD>"
     Next
     Response.Write Row & "</TR>"
    ''' rs.Delete
     rs.MoveNext	' 移到下一筆
  Wend
%>

</TABLE></CENTER>
</BODY>
</HTML>
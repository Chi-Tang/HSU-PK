<html>


<BODY>

<HEAD>

<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>�s�W����1</title>
  <script language="JavaScript">
      /// Set conn = GetMdbConnection("Test1.mdb")
 var oConn = new ActiveXObject("ADODB.Connection");
 var oRs = new ActiveXObject("ADODB.Recordset");
 
 oConn.connectionstring="DNS=TestAc-1.mdb;UID=;PWD=;"; 
 oCnn.open;
 oRs.ActiveConnection =oConn;
 oRs.Source="SeLect * From ���Z��";
 oRs.Open ;
  //if(!oRs.EOF&& !oRs.BOF){
  if(!oRs.EOF){
  DOCUMENT.WRITE("�Ǹ�: "+oRs.Fields.Item("�Ǹ�").value+<br/>);
   }
 oRs.Close;
 oConn.Close;
   </script> 


</HEAD>
</BODY>



</html>

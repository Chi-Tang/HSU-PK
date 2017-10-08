<html>

<HEAD>

<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>新增網頁1</title>
<SCRIPT LANGUAGE="VBScript">
<!--
Function CreateVBArray()
  Dim i, j, k
  Dim a(2, 2)
  k = 1
  For i = 0 To 2
    For j = 0 To 2
      a(i ,j) = k
      k = k + 1
    Next
  Next
  CreateVBArray = a
End Function
-->
</SCRIPT>
/////////////////////////////////
<SCRIPT LANGUAGE="JScript">
<!--
function VBArrayTest(vbarray)
{
  var a = new VBArray(vbarray);
  var b = a.toArray();
  var i;
  for (i = 0; i < 9; i++) 
  {
    document.writeln(b[i]);
  }
}
-->
</SCRIPT>

</HEAD>

<BODY>
<SCRIPT language="jscript">
  document.write(VBArrayTest(CreateVBArray()));
  document.write((&h5)and(&h1));
</SCRIPT>
</BODY>
</html>

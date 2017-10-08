<html>
<BODY>

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
      a(i, j) = k
      document.writeln(k)
      k = k + 1
    Next
    document.writeln("<BR>")
  Next
  CreateVBArray = a
End Function
-->
</SCRIPT>
<SCRIPT LANGUAGE="JScript">
<!--
function GetItemTest(vbarray)
{
  var i, j;
  var a = new VBArray(vbarray);
  for (i = 0; i <= 2; i++)
  {
    for (j =0; j <= 2; j++)
    {
      document.writeln(a.getItem(i, j));
      document.write((&h5)and(&h1));
    }
  }
}-->
</SCRIPT>

</HEAD>
<SCRIPT LANGUAGE="JScript">
<!--
  GetItemTest(CreateVBArray());
  
-->
</SCRIPT>
</BODY>



</html>

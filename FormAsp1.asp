<HTML>
<HEAD><TITLE>表單與 ASP 程式的合體</TITLE></HEAD>
<BODY BGCOLOR=#FFFFFF>
<H2 ALIGN=CENTER>表單與 ASP 程式的合體<HR></H2>
<% If Request("Send") = Empty Then %>
   <FORM Action=Formasp1.asp Method=POST>
   <INPUT Type=Hidden Name=Send Value="Send">
   姓名：<INPUT Type=Text Name=Name Size=20><P>
   興趣：<INPUT Type=Text Name=Love Size=20><P>
   <INPUT Type=Submit Value="傳 送">
   </FORM>
<% Else %>
   <Center><H2>
   <%=Request("Name")%> 歡迎您，您喜歡的電腦學科是
   <%=Request("Love")%> !
   <HR></H2></Center>
<% End If %>
</BODY>
</HTML>

<HTML>
<HEAD><TITLE>���P ASP �{�����X��</TITLE></HEAD>
<BODY BGCOLOR=#FFFFFF>
<H2 ALIGN=CENTER>���P ASP �{�����X��<HR></H2>
<% If Request("Send") = Empty Then %>
   <FORM Action=Formasp1.asp Method=POST>
   <INPUT Type=Hidden Name=Send Value="Send">
   �m�W�G<INPUT Type=Text Name=Name Size=20><P>
   ����G<INPUT Type=Text Name=Love Size=20><P>
   <INPUT Type=Submit Value="�� �e">
   </FORM>
<% Else %>
   <Center><H2>
   <%=Request("Name")%> �w��z�A�z���w���q���Ǭ�O
   <%=Request("Love")%> !
   <HR></H2></Center>
<% End If %>
</BODY>
</HTML>

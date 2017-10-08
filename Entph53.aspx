 
 
<Html>
<Body bgcolor="White">
<h2 align="center">@徐師易經陽宅程式@<h5 style="color: red">(用生年卦位流年紫白九星)</h5></h2>
<Form runat="server">
  <Table border="1" cellpadding="2" cellspacing="0">
    <tr>
      <td>姓名：</td>
      <td><asp:TextBox runat="server" size="20" id="Name" /></td>
    </tr>
    <tr>
      <td>性別：</td>
      <td><asp:ListBox runat="server" id="Sex" size="1" Rows=2>
        <asp:ListItem>男</asp:ListItem>
        <asp:ListItem>女</asp:ListItem>
        </asp:ListBox></td>
    </tr>
    <tr>
      <td valign="top">生年份：</td>
      <td valign="top">
      <asp:ListBox runat="server" id="BYNUM" size="6">
        <asp:ListItem>1920</asp:ListItem><asp:ListItem>1921</asp:ListItem><asp:ListItem>1922</asp:ListItem><asp:ListItem>1923</asp:ListItem><asp:ListItem>1924</asp:ListItem>
        <asp:ListItem>1925</asp:ListItem><asp:ListItem>1926</asp:ListItem><asp:ListItem>1927</asp:ListItem><asp:ListItem>1928</asp:ListItem><asp:ListItem>1929</asp:ListItem>
        <asp:ListItem>1930</asp:ListItem><asp:ListItem>1931</asp:ListItem><asp:ListItem>1932</asp:ListItem><asp:ListItem>1933</asp:ListItem><asp:ListItem>1934</asp:ListItem>
        <asp:ListItem>1935</asp:ListItem><asp:ListItem>1936</asp:ListItem><asp:ListItem>1937</asp:ListItem><asp:ListItem>1938</asp:ListItem><asp:ListItem>1939</asp:ListItem>
        <asp:ListItem>1940</asp:ListItem><asp:ListItem>1941</asp:ListItem><asp:ListItem>1942</asp:ListItem><asp:ListItem>1943</asp:ListItem><asp:ListItem>1944</asp:ListItem>
        <asp:ListItem>1945</asp:ListItem><asp:ListItem>1946</asp:ListItem><asp:ListItem>1947</asp:ListItem><asp:ListItem>1948</asp:ListItem><asp:ListItem>1949</asp:ListItem>
        <asp:ListItem>1950</asp:ListItem><asp:ListItem>1951</asp:ListItem><asp:ListItem>1952</asp:ListItem><asp:ListItem>1953</asp:ListItem><asp:ListItem>1954</asp:ListItem>
        <asp:ListItem>1955</asp:ListItem><asp:ListItem>1956</asp:ListItem><asp:ListItem>1957</asp:ListItem><asp:ListItem>1958</asp:ListItem><asp:ListItem>1959</asp:ListItem>
        <asp:ListItem>1960</asp:ListItem><asp:ListItem>1961</asp:ListItem><asp:ListItem>1962</asp:ListItem><asp:ListItem>1963</asp:ListItem><asp:ListItem>1964</asp:ListItem>
        <asp:ListItem>1965</asp:ListItem><asp:ListItem>1966</asp:ListItem><asp:ListItem>1967</asp:ListItem><asp:ListItem>1968</asp:ListItem><asp:ListItem>1969</asp:ListItem>
        <asp:ListItem>1970</asp:ListItem><asp:ListItem>1971</asp:ListItem><asp:ListItem>1972</asp:ListItem><asp:ListItem>1973</asp:ListItem><asp:ListItem>1974</asp:ListItem>
        <asp:ListItem>1975</asp:ListItem><asp:ListItem>1976</asp:ListItem><asp:ListItem>1977</asp:ListItem><asp:ListItem>1978</asp:ListItem><asp:ListItem>1979</asp:ListItem>
        <asp:ListItem>1980</asp:ListItem><asp:ListItem>1981</asp:ListItem><asp:ListItem>1982</asp:ListItem><asp:ListItem>1983</asp:ListItem><asp:ListItem>1984</asp:ListItem>
        <asp:ListItem>1985</asp:ListItem><asp:ListItem>1986</asp:ListItem><asp:ListItem>1987</asp:ListItem><asp:ListItem>1988</asp:ListItem><asp:ListItem>1989</asp:ListItem>
        <asp:ListItem>1990</asp:ListItem><asp:ListItem>1991</asp:ListItem><asp:ListItem>1992</asp:ListItem><asp:ListItem>1993</asp:ListItem><asp:ListItem>1994</asp:ListItem>
        <asp:ListItem>1995</asp:ListItem><asp:ListItem>1996</asp:ListItem><asp:ListItem>1997</asp:ListItem><asp:ListItem>1998</asp:ListItem><asp:ListItem>1999</asp:ListItem>
        <asp:ListItem>2000</asp:ListItem><asp:ListItem>2001</asp:ListItem><asp:ListItem>2002</asp:ListItem><asp:ListItem>2003</asp:ListItem><asp:ListItem>2004</asp:ListItem>
        <asp:ListItem>2005</asp:ListItem><asp:ListItem>2006</asp:ListItem><asp:ListItem>2007</asp:ListItem><asp:ListItem>2008</asp:ListItem><asp:ListItem>2009</asp:ListItem>
        <asp:ListItem>2010</asp:ListItem><asp:ListItem>2011</asp:ListItem><asp:ListItem>2012</asp:ListItem><asp:ListItem>2013</asp:ListItem><asp:ListItem>2014</asp:ListItem>
        <asp:ListItem>2015</asp:ListItem><asp:ListItem>2016</asp:ListItem><asp:ListItem>2017</asp:ListItem><asp:ListItem>2018</asp:ListItem><asp:ListItem>2019</asp:ListItem>
        <asp:ListItem>2020</asp:ListItem><asp:ListItem>2021</asp:ListItem><asp:ListItem>2022</asp:ListItem><asp:ListItem>2023</asp:ListItem><asp:ListItem>2024</asp:ListItem>
        <asp:ListItem>2025</asp:ListItem>
      
     </asp:ListBox></td>
    </tr>
    <tr>
      <td valign="top">流年份：</td>
      <td valign="top">
      <asp:ListBox runat="server" id="LYNUM" size="6">
        <asp:ListItem>1920</asp:ListItem><asp:ListItem>1921</asp:ListItem><asp:ListItem>1922</asp:ListItem><asp:ListItem>1923</asp:ListItem><asp:ListItem>1924</asp:ListItem>
        <asp:ListItem>1925</asp:ListItem><asp:ListItem>1926</asp:ListItem><asp:ListItem>1927</asp:ListItem><asp:ListItem>1928</asp:ListItem><asp:ListItem>1929</asp:ListItem>
        <asp:ListItem>1930</asp:ListItem><asp:ListItem>1931</asp:ListItem><asp:ListItem>1932</asp:ListItem><asp:ListItem>1933</asp:ListItem><asp:ListItem>1934</asp:ListItem>
        <asp:ListItem>1935</asp:ListItem><asp:ListItem>1936</asp:ListItem><asp:ListItem>1937</asp:ListItem><asp:ListItem>1938</asp:ListItem><asp:ListItem>1939</asp:ListItem>
        <asp:ListItem>1940</asp:ListItem><asp:ListItem>1941</asp:ListItem><asp:ListItem>1942</asp:ListItem><asp:ListItem>1943</asp:ListItem><asp:ListItem>1944</asp:ListItem>
        <asp:ListItem>1945</asp:ListItem><asp:ListItem>1946</asp:ListItem><asp:ListItem>1947</asp:ListItem><asp:ListItem>1948</asp:ListItem><asp:ListItem>1949</asp:ListItem>
        <asp:ListItem>1950</asp:ListItem><asp:ListItem>1951</asp:ListItem><asp:ListItem>1952</asp:ListItem><asp:ListItem>1953</asp:ListItem><asp:ListItem>1954</asp:ListItem>
        <asp:ListItem>1955</asp:ListItem><asp:ListItem>1956</asp:ListItem><asp:ListItem>1957</asp:ListItem><asp:ListItem>1958</asp:ListItem><asp:ListItem>1959</asp:ListItem>
        <asp:ListItem>1960</asp:ListItem><asp:ListItem>1961</asp:ListItem><asp:ListItem>1962</asp:ListItem><asp:ListItem>1963</asp:ListItem><asp:ListItem>1964</asp:ListItem>
        <asp:ListItem>1965</asp:ListItem><asp:ListItem>1966</asp:ListItem><asp:ListItem>1967</asp:ListItem><asp:ListItem>1968</asp:ListItem><asp:ListItem>1969</asp:ListItem>
        <asp:ListItem>1970</asp:ListItem><asp:ListItem>1971</asp:ListItem><asp:ListItem>1972</asp:ListItem><asp:ListItem>1973</asp:ListItem><asp:ListItem>1974</asp:ListItem>
        <asp:ListItem>1975</asp:ListItem><asp:ListItem>1976</asp:ListItem><asp:ListItem>1977</asp:ListItem><asp:ListItem>1978</asp:ListItem><asp:ListItem>1979</asp:ListItem>
        <asp:ListItem>1980</asp:ListItem><asp:ListItem>1981</asp:ListItem><asp:ListItem>1982</asp:ListItem><asp:ListItem>1983</asp:ListItem><asp:ListItem>1984</asp:ListItem>
        <asp:ListItem>1985</asp:ListItem><asp:ListItem>1986</asp:ListItem><asp:ListItem>1987</asp:ListItem><asp:ListItem>1988</asp:ListItem><asp:ListItem>1989</asp:ListItem>
        <asp:ListItem>1990</asp:ListItem><asp:ListItem>1991</asp:ListItem><asp:ListItem>1992</asp:ListItem><asp:ListItem>1993</asp:ListItem><asp:ListItem>1994</asp:ListItem>
        <asp:ListItem>1995</asp:ListItem><asp:ListItem>1996</asp:ListItem><asp:ListItem>1997</asp:ListItem><asp:ListItem>1998</asp:ListItem><asp:ListItem>1999</asp:ListItem>
        <asp:ListItem>2000</asp:ListItem><asp:ListItem>2001</asp:ListItem><asp:ListItem>2002</asp:ListItem><asp:ListItem>2003</asp:ListItem><asp:ListItem>2004</asp:ListItem>
        <asp:ListItem>2005</asp:ListItem><asp:ListItem>2006</asp:ListItem><asp:ListItem>2007</asp:ListItem><asp:ListItem>2008</asp:ListItem><asp:ListItem>2009</asp:ListItem>
        <asp:ListItem>2010</asp:ListItem><asp:ListItem>2011</asp:ListItem><asp:ListItem>2012</asp:ListItem><asp:ListItem>2013</asp:ListItem><asp:ListItem>2014</asp:ListItem>
        <asp:ListItem>2015</asp:ListItem><asp:ListItem>2016</asp:ListItem><asp:ListItem>2017</asp:ListItem><asp:ListItem>2018</asp:ListItem><asp:ListItem>2019</asp:ListItem>
        <asp:ListItem>2020</asp:ListItem><asp:ListItem>2021</asp:ListItem><asp:ListItem>2022</asp:ListItem><asp:ListItem>2023</asp:ListItem><asp:ListItem>2024</asp:ListItem>
        <asp:ListItem>2025</asp:ListItem>
      
     </asp:ListBox></td>
    </tr>
    <tr>
      <td valign="top">公司(機關)宅坐山卦：</td>
      <td valign="top">
      <asp:ListBox runat="server" id="HSNUM" size="6">
        <asp:ListItem>坐坎向離</asp:ListItem><asp:ListItem>坐坤向艮</asp:ListItem><asp:ListItem>坐震向兌</asp:ListItem><asp:ListItem>坐巽向乾</asp:ListItem>
       <asp:ListItem>坐乾向巽</asp:ListItem><asp:ListItem>坐兌向震</asp:ListItem><asp:ListItem>坐艮向坤</asp:ListItem><asp:ListItem>坐離向坎</asp:ListItem>
          </asp:ListBox></td>
    </tr>
   <tr>
      <td valign="top">三元九大運：</td>
      <td valign="top">
      <asp:ListBox runat="server" id="LNNUM" size="6">
        <asp:ListItem>1</asp:ListItem><asp:ListItem>2</asp:ListItem><asp:ListItem>3</asp:ListItem><asp:ListItem>4</asp:ListItem><asp:ListItem>5</asp:ListItem>
       <asp:ListItem>6</asp:ListItem><asp:ListItem>7</asp:ListItem><asp:ListItem>8</asp:ListItem><asp:ListItem>9</asp:ListItem>
          </asp:ListBox></td>
    </tr>
  </Table>
  <asp:Button runat="server" Text=" 傳 送 " OnClick="Button_Click" />
</Form>
</Body>
</Html>

<script Language="VB" runat="server">
   Sub Button_Click(sender As Object, e As EventArgs) 
       Dim URL, LoveURL, I
        ''Dim YNUM, MNUM, DNUM, HNUM
       URL  = "Excph53.aspx"
       URL &= "?Name=" & Server.URLEncode( Name.Text )
       URL &= "&Sex=" & Server.URLEncode( Sex.SelectedItem.Text )
       URL &= "&BYNUM=" & Server.URLEncode( BYNUM.SelectedItem.Text )
       URL &= "&LYNUM=" & Server.URLEncode( LYNUM.SelectedItem.Text )
        URL &= "&LNNUM=" & Server.UrlEncode(LNNUM.SelectedItem.Text)
        URL &= "&HSNUM=" & Server.UrlEncode(HSNUM.SelectedItem.Text)
       Server.Transfer( URL )
   End Sub
</script>

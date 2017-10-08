 
 
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
 <asp:ListItem>壬山丙向正坐向</asp:ListItem><asp:ListItem>壬山丙向兼子午</asp:ListItem><asp:ListItem>壬山丙向兼亥巳</asp:ListItem>
<asp:ListItem>子山午向正坐向</asp:ListItem><asp:ListItem>子山午向兼壬丙</asp:ListItem><asp:ListItem>子山午向兼癸丁</asp:ListItem>
<asp:ListItem>癸山丁向正坐向</asp:ListItem><asp:ListItem>癸山丁向兼子午</asp:ListItem><asp:ListItem>癸山丁向兼丑未</asp:ListItem> 
<asp:ListItem>丑山未向正坐向</asp:ListItem><asp:ListItem>丑山未向兼癸丁</asp:ListItem><asp:ListItem>丑山未向兼艮坤</asp:ListItem> 
<asp:ListItem>艮山坤向正坐向</asp:ListItem><asp:ListItem>艮山坤向兼丑未</asp:ListItem><asp:ListItem>艮山坤向兼寅申</asp:ListItem> 
<asp:ListItem>寅山申向正坐向</asp:ListItem><asp:ListItem>寅山申向兼艮坤</asp:ListItem><asp:ListItem>寅山申向兼甲庚</asp:ListItem> 
<asp:ListItem>甲山庚向正坐向</asp:ListItem><asp:ListItem>甲山庚向兼寅申</asp:ListItem><asp:ListItem>甲山庚向兼卯酉</asp:ListItem> 
<asp:ListItem>卯山酉向正坐向</asp:ListItem><asp:ListItem>卯山酉向兼甲庚</asp:ListItem><asp:ListItem>卯山酉向兼乙辛</asp:ListItem> 
<asp:ListItem>乙山辛向正坐向</asp:ListItem><asp:ListItem>乙山辛向兼卯酉</asp:ListItem><asp:ListItem>乙山辛向兼辰戌</asp:ListItem> 
<asp:ListItem>辰山戌向正坐向</asp:ListItem><asp:ListItem>辰山戌向兼乙辛</asp:ListItem><asp:ListItem>辰山戌向兼巽乾</asp:ListItem> 
<asp:ListItem>巽山乾向正坐向</asp:ListItem><asp:ListItem>巽山乾向兼辰戌</asp:ListItem><asp:ListItem>巽山乾向兼巳亥</asp:ListItem> 
<asp:ListItem>巳山亥向正坐向</asp:ListItem><asp:ListItem>巳山亥向兼巽乾</asp:ListItem><asp:ListItem>巳山亥向兼丙壬</asp:ListItem> 
<asp:ListItem>丙山壬向正坐向</asp:ListItem><asp:ListItem>丙山壬向兼巳亥</asp:ListItem><asp:ListItem>丙山壬向兼午子</asp:ListItem> 
<asp:ListItem>午山子向正坐向</asp:ListItem><asp:ListItem>午山子向兼丙壬</asp:ListItem><asp:ListItem>午山子向兼丁癸</asp:ListItem> 
<asp:ListItem>丁山癸向正坐向</asp:ListItem><asp:ListItem>丁山癸向兼午子</asp:ListItem><asp:ListItem>丁山癸向兼未丑</asp:ListItem> 
<asp:ListItem>未山丑向正坐向</asp:ListItem><asp:ListItem>未山丑向兼丁癸</asp:ListItem><asp:ListItem>未山丑向兼坤艮</asp:ListItem>
<asp:ListItem>坤山艮向正坐向</asp:ListItem><asp:ListItem>坤山艮向兼丑未</asp:ListItem><asp:ListItem>坤山艮向兼申寅</asp:ListItem> 
<asp:ListItem>申山寅向正坐向</asp:ListItem><asp:ListItem>申山寅向兼坤艮</asp:ListItem><asp:ListItem>申山寅向兼庚甲</asp:ListItem> 
<asp:ListItem>庚山甲向正坐向</asp:ListItem><asp:ListItem>庚山甲向兼甲寅</asp:ListItem><asp:ListItem>庚山甲向兼酉卯</asp:ListItem> 
<asp:ListItem>酉山卯向正坐向</asp:ListItem><asp:ListItem>酉山卯向兼庚甲</asp:ListItem><asp:ListItem>酉山卯向兼辛乙</asp:ListItem> 
<asp:ListItem>辛山乙向正坐向</asp:ListItem><asp:ListItem>辛山乙向兼酉卯</asp:ListItem><asp:ListItem>辛山乙向兼戌辰</asp:ListItem> 
<asp:ListItem>戌山辰向正坐向</asp:ListItem><asp:ListItem>戌山辰向兼辛乙</asp:ListItem><asp:ListItem>戌山辰向兼乾巽</asp:ListItem> 
<asp:ListItem>乾山巽向正坐向</asp:ListItem><asp:ListItem>乾山巽向兼戌辰</asp:ListItem><asp:ListItem>乾山巽向兼亥巳</asp:ListItem>
<asp:ListItem>亥山巳向正坐向</asp:ListItem><asp:ListItem>亥山巳向兼乾巽</asp:ListItem><asp:ListItem>亥山巳向兼壬丙</asp:ListItem> 
  </asp:ListBox></td>
    </tr>
   <tr>
      <td valign="top">三元九大運：</td>
      <td valign="top">
      <asp:ListBox runat="server" id="LNNUM" size="6">
        <asp:ListItem>1</asp:ListItem><asp:ListItem>2</asp:ListItem><asp:ListItem>3</asp:ListItem><asp:ListItem>4</asp:ListItem><asp:ListItem>5</asp:ListItem>
       <asp:ListItem>6</asp:ListItem><asp:ListItem>7</asp:ListItem><asp:ListItem>8</asp:ListItem><asp:ListItem>9</asp:ListItem>
          </asp:ListBox>
      </td>
    </tr>
  </Table>
  <asp:Button runat="server" Text=" 傳 送 " OnClick="Button_Click" />
</Form>
</Body>
</Html>

<script Language="VB" runat="server">
    Dim TNTFS() = {"0", "壬", "子", "癸", "丑", "艮", "寅", "甲", "卯", "乙", "辰", "巽", "巳", "丙", "午", "丁", "未", "坤", "申", "庚", "酉", "辛", "戌", "乾", "亥"}
    Sub Button_Click(sender As Object, e As EventArgs)
        Dim URL, LoveURL, I
        ''Dim YNUM, MNUM, DNUM, HNUM
        URL = "Excph515.aspx"
        URL &= "?Name=" & Server.UrlEncode(Name.Text)
        URL &= "&Sex=" & Server.UrlEncode(Sex.SelectedItem.Text)
        URL &= "&BYNUM=" & Server.UrlEncode(BYNUM.SelectedItem.Text)
        URL &= "&LYNUM=" & Server.UrlEncode(LYNUM.SelectedItem.Text)
        URL &= "&LNNUM=" & Server.UrlEncode(LNNUM.SelectedItem.Text)
        URL &= "&HSNUM=" & Server.UrlEncode(HSNUM.SelectedItem.Text)
        Server.Transfer(URL)
    End Sub
</script>

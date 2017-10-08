 
 <%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Web.Mail" %>

 <%
  
     Dim TNam, TNsx, TNhs, InpLn, InpSFx, InpSx, InpFx
  Dim TNbr, TNyr,Nob, Noy ,BNum, YNum ,RTNBY ,RBYNS,RBN, RYN

  TNam = Request("Name")
  TNsx = Request("Sex")
     TNbr = Request("BYNUM")
     TNyr = Request("LYNUM")
     InpLn = Request("LNNUM")
     InpSFx = Request("HSNUM")
     ''TNhs=Request("HSNUM")
     TNhs = Mid(InpSFx, 2, 1)
     InpSx = Mid(InpSFx, 2, 1)
     InpFx = Mid(InpSFx, 4, 1)
     Nob = CInt(TNbr)
     Noy = CInt(TNyr)
     Dim R() = {" ", "[坎宮]", "[坤宮]", "[震宮]", "[巽宮]", "[中宮]", "[乾宮]", "[兌宮]", "[艮宮]", "[離宮]", "[黃宮]"}
     Dim ZNAM() = {"離", "坎", "坤", "震", "巽", "黃", "乾", "兌", "艮", "離"}
     Dim ZNUM() = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}
     Dim I, J, k, n, BBK As Integer
     ''Dim ii, isn, ifn 
      On Error Resume Next
     Dim A(60), B(60)
     k = 30
     Dim ASTR1 = STRLSF(InpLn)
     Dim ASTR2 = STRSS(InpLn, InpSx)
     Dim  ASTR3 = STRSS(InpLn, InpFx)
     '' Dim ASTR3 = STRSS(InpLn, InpFn, "正兼")
     For n = 1 To 9
         B(30 + n) = n
         A(30 + n) = ASTR1(n) & "L"
         B(40 + n) = n
         A(40 + n) = ASTR2(n) & "S"
         B(50 + n) = n
         A(50 + n) = ASTR3(n) & "F"
     Next
     Select Case TNsx
         Case "男"
             If Nob >= 1863 Then
                 BNum = (11 - ((Nob - 1863) Mod 9)) Mod 9
                 '' Else
                 '' Response.Write("值年超上元範圍")
             End If
             If BNum = 5 Then
                 BNum = 2
             End If
         Case "女"
             If Nob >= 1863 Then
                 BNum = (4 + ((Nob - 1863) Mod 9)) Mod 9
                 ''Else
                 '' Response.Write("值年超上元範圍")
             End If
             If BNum = 5 Then
                 BNum = 8
             End If
     End Select
     If BNum = 0 Then
         BNum = 9
     End If
     
    
     If Noy >= 1863 Then
         YNum = (11 - ((Noy - 1863) Mod 9)) Mod 9
         ''Else
        '' Response.Write("值年超上元範圍")
     End If
     '' If YNum = 5 Then  ''八宅生(流)年卦才使用"
     ''Select Case TNsx
     ''    Case "男"
     '' YNum = 2
     ''     Case "女"
     ''  YNum = 8
     ''  End Select
     '''  End If
     If YNum = 0 Then
         YNum = 9
     End If

     Response.Write(BNum & ";" & YNum & "<br>")

     '' RTNBY = QueryDataAndSendTo(BNum, YNum)
     '' RBYNS = Split( RTNBY ,";")
     '  RBN =RBYNS
     '  RYN =RBYNS

     ''''''''''''''''''''''''''''''''''''''''''''''''''''''
     '''' <script Language="VB" runat="server">
     '''' Function  QueryDataAndSendTo(XNam1,XNam2)
    
     On Error Resume Next
     Dim Conn As OleDbConnection, Cmd As OleDbCommand
     Dim Rdb, Rdy, Rhs As OleDbDataReader, SQL As String, Body As String
     Dim mad11, mad22 As String
    '' Dim I, J, k, BBK As Integer
    '' Dim A(30), B(30)
     k = 0
     Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      
     ''Dim Database = "Data Source=" & Server.MapPath( "../ch15/Users.mdb" )
     'Dim Database = "Data Source=" & Server.MapPath( "UsersPwd.mdb" )
     'Dim Database = "Data Source=" & Server.MapPath( "/HSU-fundb/UsersPwd.mdb" )
     Dim Database = "Data Source=" & Server.MapPath("/HSU-PK/UsersPwd.mdb")
     Dim Dbpass = "Jet OLEDB:Database Password=kj6688"
     Conn = New OleDbConnection(Provider & ";" & Database & ";" & Dbpass)
     Conn.Open()
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
     Dim Zdb(9)
     Dim Zdbk = " "
     k = 1
     ' 檢查 Email 是否存在
     ' SQL = "Select * From Users Where Email='" & Emad1.Text & "'"
     'SQL = "Select * From 八卦表 Where 內卦='" & Emad1.Text & "'"
     SQL = "Select * From 生年卦位表"
     Cmd = New OleDbCommand(SQL, Conn)
     Rdb = Cmd.ExecuteReader()
     While Rdb.Read()
         If Rdb.Item(0) = BNum Then ' 表示此一 Email 存在
             'If Rdb.Item("內卦")=Emad1.Text Then ' 表示此一 Email 存在
             'If Rdb .Read() Then ' 表示此一 Email 存在
             ' Dim mad1 As string
             For I = 1 To Rdb.FieldCount - 1
                 '' mad11 = Rdb.Item("天")
                 '' if Rdb .GetName(i) = XNam2 then
                 ''mad22 =Rdb.Item(i)
                 ' Zdb(i)=Rdb.GetName(i)
                 ' Zdbk =Rdb.Item(i)

                 A(k + I) = "(" & Rdb.GetName(I) & ")"
                 B(k + I) = Rdb.Item(I)
 
                 '' End If 

                 '' Response.Write (A(k+i)& B(k+i))
             Next
         End If
         'A(0+k)=Zdb(i)
         'B(0+k)=Zdbk
         ''k=k+1
     End While
     Rdb.Close()
     '' Conn.Close()
  
     Dim Zdy(9)
     Dim Zdyk = " "
     k = 10
     ' 檢查 Email 是否存在
     ' SQL = "Select * From Users Where Email='" & Emad1.Text & "'"
     'SQL = "Select * From 八卦表 Where 內卦='" & Emad1.Text & "'"
     SQL = "Select * From 流年紫白表"
     Cmd = New OleDbCommand(SQL, Conn)
     Rdy = Cmd.ExecuteReader()
     While Rdy.Read()
         If Rdy.Item(0) = YNum Then ' 表示此一 Email 存在
             'If Rdy.Item("內卦")=Emad1.Text Then ' 表示此一 Email 存在
             'If Rdy .Read() Then ' 表示此一 Email 存在
             ' Dim mad1 As string
             For I = 1 To Rdy.FieldCount - 1
                 '' mad11 = Rdy.Item("天")
                 '' if Rdy .GetName(i) = XNam2 then
                 ''mad22 =Rdy.Item(i)
                 ' Zdy(i)=Rdy.GetName(i)
                 '  Zdyk =Rdy.Item(i)

                 A(k + I) = Rdy.GetName(I)
                 B(k + I) = Rdy.Item(I)
                 If Rdy.Item(I) = 5 Then
                     A(k + I) = "流年星:" & A(k + I)
                 End If
                 '' End If 
                 '' Response.Write (A(k+i)& B(k+i))
             Next
         End If
         ''A(0+k)=Zdy(i)
         '' B(0+k)=Zdyk
  
         ''k=k+1
 
     End While
     Rdy.Close()
     '' Conn.Close()
   
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     Dim Zhs(9)
     Dim Zhsk = " "
     k = 20
     ' 檢查 Email 是否存在
     ' SQL = "Select * From Users Where Email='" & Emad1.Text & "'"
     'SQL = "Select * From 八卦表 Where 內卦='" & Emad1.Text & "'"
     SQL = "Select * From 宅坐飛卦表"
     '' SQL = "Select * From 宅坐紫白表"
     Cmd = New OleDbCommand(SQL, Conn)
     Rhs = Cmd.ExecuteReader()
     While Rhs.Read()
         If Rhs.Item(0) = TNhs Then ' 表示此一 Email 存在
             'If Rhs.Item("內卦")=Emad1.Text Then ' 表示此一 Email 存在
             'If Rhs .Read() Then ' 表示此一 Email 存在
             ' Dim mad1 As string
             For I = 1 To Rhs.FieldCount - 1
                 '' mad11 = Rhs.Item("天")
                 '' if Rhs .GetName(i) = XNam2 then
                 ''mad22 =Rhs.Item(i)
                 B(k + I) = Rhs.GetName(I)
                 A(k + I) = Rhs.Item(I)
                 If Rhs.GetName(I) = 5 Then
                     A(k + I) = "坐山星:" & A(k + I)
                 End If
                 '' A(k + I) = Rhs.GetName(I) ''B()存卦數,A()存項目
                 '' B(k + I) = Rhs.Item(I)   '' 宅坐飛卦表與紫白表互調 
                 '' End If 
                 '' Response.Write (A(k+i)& B(k+i))
             Next
         End If
         ''A(0+k)=Zhs(i)
         '' B(0+k)=Zhsk
  
         ''k=k+1
 
     End While
     Rhs.Close()
     Conn.Close()

     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
     For J = 0 To 60
         BBK = Trim(B(J))
         Select Case BBK
             Case "1"
                 ''R(1)=Trim(R(1))+","& Trim(A(j))
                 R(1) = Trim(R(1)) + "<br>" & Trim(A(J))
             Case "2"
                 R(2) = Trim(R(2)) & "<br>" & Trim(A(J))
             Case "3"
                 R(3) = Trim(R(3)) + "<br>" & Trim(A(J))
             Case "4"
                 R(4) = Trim(R(4)) + "<br>" & Trim(A(J))
             Case "5"
                 R(5) = Trim(R(5)) + "<br>" & Trim(A(J))
             Case "6"
                 R(6) = Trim(R(6)) + "<br>" & Trim(A(J))
             Case "7"
                 R(7) = Trim(R(7)) + "<br>" & Trim(A(J))
             Case "8"
                 R(8) = Trim(R(8)) + "<br>" & Trim(A(J))
             Case "9"
                 R(9) = Trim(R(9)) + "<br>" & Trim(A(J))
             Case Else
                 R(10) = Trim(R(10)) + "<br>" & Trim(A(J))
         End Select
     Next

  
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''      
     Dim Row1 = "<CENTER><TABLE Border=1 BGCOLOR=#FFFF00 >" & "<TR>" & "<TD width=120>卦名" & "</TD>" & "<TD width=120>(生成卦)</TD>" & "<TD width =120>流之卦</TD>" & "</TR>"
     Dim Row2 = "<TR>" & "<TD width=120>" & R(4) & "</TD>" & "<TD width=120>" & R(9) & "</TD>" & "<TD width =120>" & R(2) & "</TD>" & "</TR>"
     Dim Row3 = "<TR>" & "<TD width=120>" & R(3) & "</TD>" & "<TD width=120>" & R(5) & "</TD>" & "<TD width =120>" & R(7) & "</TD>" & "</TR>"
     Dim Row4 = "<TR>" & "<TD width=120>" & R(8) & "</TD>" & "<TD width=120>" & R(1) & "</TD>" & "<TD width =120>" & R(6) & "</TD>" & "</TR>"
     Dim Row5 = "<TR style='writing-mode:bt-rl' width =150>" & "尊重著作權:<u>免費分享</u>" & "</TR>" & "</TABLE>"
     Dim RowT = Row1 + Row2 + Row3 + Row4 + Row5
     Response.Write(RowT)
 
 %>
<script Language="VB" runat="server">
   
    Dim ZNAM() = {"離", "坎", "坤", "震", "巽", "黃", "乾", "兌", "艮", "離"}
    Dim ZNUM() = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}
    ''Dim ZNFS()={"午","子","坤","卯","巽","中","乾","酉","艮","午"}  
    ''Dim TNTFS() = {"0", "壬", "子", "癸", "丑", "艮", "寅", "甲", "卯", "乙", "辰", "巽", "巳", "丙", "午", "丁", "未", "坤", "申", "庚", "酉", "辛", "戌", "乾", "亥"}
    Dim TNTFS() = {"0", "壬", "坎", "癸", "丑", "艮", "寅", "甲", "震", "乙", "辰", "巽", "巳", "丙", "離", "丁", "未", "坤", "申", "庚", "兌", "辛", "戌", "乾", "亥"}
    Dim TNTFN() = {"0", "11", "12", "13", "81", "82", "83", "31", "32", "33", "41", "42", "43", "91", "92", "93", "21", "22", "23", "71", "72", "73", "61", "62", "63"}
    Dim TNTFX() = {"0", "陽", "陰", "陰", "陰", "陽", "陽", "陽", "陰", "陰", "陰", "陽", "陽", "陽", "陰", "陰", "陰", "陽", "陽", "陽", "陰", "陰", "陰", "陽", "陽"}
   
    Function STRLSF(LN)
        
        Dim i, j, k, m, n, xo, yo, tk, pkn As Integer
        Dim APK(9), ATR(9)
        '' Dim LK(10), LLUK(10), LR(60), LYER(60)
        ''Dim TDAYS = Split(SDAY, ":")
        ''Dim TDAYS1 = CInt(TDAYS(0))
        '' Dim TDAYS2 = CInt(TDAYS(1))
        For i = 1 To 9
            If ZNUM(i) = CInt(LN) Then
                tk = i
                Exit For
            End If
        Next
       
        For j = 1 To 9
            pkn = ((i + 4) + j + 9) Mod 9
            If pkn = 0 Then
                pkn = 9
            End If
          
            APK(j) = j
            ATR(j) = ZNUM(pkn)
        Next
        Return ATR
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '' Response.Write(ATR(8))
    End Function
    Function STRLSB(LN)
       
        Dim i, j, k, m, n, xo, yo, tk As Integer
        Dim APK(9), ATR(9)
        
        For i = 1 To 9
            If ZNUM(i) = CInt(LN) Then
                tk = i
                Exit For
            End If
        Next
       
        For j = 1 To 9
            Dim pkn = ((i + 5) - j + 9) Mod 9
            If pkn = 0 Then
                pkn = 9
            End If
          
            APK(j) = j
            ATR(j) = ZNUM(pkn)
        Next
        Return ATR
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''Response.Write(ATR(6)
    End Function
    
    Function STRSS(LN, SX)
      
        Dim i, j, k, m, n, tkpn, tksn, gtkn
        Dim sexy, xy, tkps
        Dim APK(9), ATR1(9)
        '' Dim LK(10), LLUK(10), LR(60), LYER(60)
        ''Dim TDAYS = Split(SDAY, ":")
        ''Dim TDAYS1 = CInt(TDAYS(0))
        '' Dim TDAYS2 = CInt(TDAYS(1))
        For i = 1 To 24
            If TNTFS(i) = Trim(SX) Then
                tkps = TNTFN(i)
                sexy = TNTFX(i)
                Exit For
            End If
        Next
        tkpn = CInt(Left(tkps, 1))
        tksn = CInt(Right(tkps, 1))
        Dim ASTR = STRLSF(LN)
        ' For j = 1 To 9
        'If j = tkpn Then
        ''gtks = ASTR(j
        gtkn = ASTR(tkpn)
        ' Exit For
        ' End If
        ' Next
        If gtkn <> 5 Then
            For k = 1 To 24
                If TNTFN(k) = Trim(gtkn & tksn) Then
                    xy = TNTFX(k)
                    Exit For
                End If
            Next
        Else
            '' gtkn = 5
            xy = Trim(sexy)
        End If
        Select Case xy
            Case "陽"
                ATR1 = STRLSF(gtkn)
            Case "陰"
                ATR1 = STRLSB(gtkn)
                          
        End Select
        Return ATR1
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '' Response.Write(tkps & xy & ATR1(5))
    End Function
</script>
  <HTML> 
  <BODY bgcolor="#FFFFFF">
  <CENTER><H2> 易 經 八 卦 盤 與 評 論 表 如 下<HR></H2>
  <TABLE Border=1 Width=60%  BGCOLOR=#FFFF00>
  
  <%  
 ''Dim Row1 ,Row2 ,Row3 ,Row4 ,Row5 ,Row6 ,Row7 ,Row8 ,RowT
      ''    Response.Write (RowT)
       
  %>

</TABLE></CENTER>
 </BODY>
 
  <!-- <CENTER><H2><FONT color=#FF0000>若需要<u>卦 盤 評 論 表</u></FONT><HR></H2>
 <CENTER><H2><FONT color=#FF0000>    請 Email:  tech.t1206@gmail.com 即可聯絡,  謝謝</FONT><HR></H2>
  -->
</Html>
 
 
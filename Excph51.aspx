 
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
     Dim R() = {" ", "[§¢®c]", "[©[®c]", "[¾_®c]", "[´S®c]", "[¤¤®c]", "[°®®c]", "[§I®c]", "[¦á®c]", "[Â÷®c]", "[¶À®c]"}
     Dim ZNAM() = {"Â÷", "§¢", "©[", "¾_", "´S", "¶À", "°®", "§I", "¦á", "Â÷"}
     Dim ZNUM() = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}
     Dim I, J, k, n, BBK As Integer
     ''Dim ii, isn, ifn 
      On Error Resume Next
     Dim A(60), B(60)
     k = 30
     Dim ASTR1 = STRLSF(InpLn)
     Dim ASTR2 = STRSS(InpLn, InpSx)
     Dim  ASTR3 = STRSS(InpLn, InpFx)
     '' Dim ASTR3 = STRSS(InpLn, InpFn, "¥¿­Ý")
     For n = 1 To 9
         B(30 + n) = n
         A(30 + n) = ASTR1(n) & "L"
         B(40 + n) = n
         A(40 + n) = ASTR2(n) & "S"
         B(50 + n) = n
         A(50 + n) = ASTR3(n) & "F"
     Next
     
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
     Dim Row1 = "<CENTER><TABLE Border=1 BGCOLOR=#FFFF00 >" & "<TR>" & "<TD width=120>¨ö¦W" & "</TD>" & "<TD width=120>(¥Í¦¨¨ö)</TD>" & "<TD width =120>¬y¤§¨ö</TD>" & "</TR>"
     Dim Row2 = "<TR>" & "<TD width=120>" & R(4) & "</TD>" & "<TD width=120>" & R(9) & "</TD>" & "<TD width =120>" & R(2) & "</TD>" & "</TR>"
     Dim Row3 = "<TR>" & "<TD width=120>" & R(3) & "</TD>" & "<TD width=120>" & R(5) & "</TD>" & "<TD width =120>" & R(7) & "</TD>" & "</TR>"
     Dim Row4 = "<TR>" & "<TD width=120>" & R(8) & "</TD>" & "<TD width=120>" & R(1) & "</TD>" & "<TD width =120>" & R(6) & "</TD>" & "</TR>"
     Dim Row5 = "<TR style='writing-mode:bt-rl' width =150>" & "´L­«µÛ§@Åv:<u>§K¶O¤À¨É</u>" & "</TR>" & "</TABLE>"
     Dim RowT = Row1 + Row2 + Row3 + Row4 + Row5
     Response.Write(RowT)
 
 %>
<script Language="VB" runat="server">
   
    Dim ZNAM() = {"Â÷", "§¢", "©[", "¾_", "´S", "¶À", "°®", "§I", "¦á", "Â÷"}
    Dim ZNUM() = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}
    ''Dim ZNFS()={"¤È","¤l","©[","¥f","´S","¤¤","°®","¨»","¦á","¤È"}  
    ''Dim TNTFS() = {"0", "¤Ð", "¤l", "¬Ñ", "¤¡", "¦á", "±G", "¥Ò", "¥f", "¤A", "¨°", "´S", "¤x", "¤þ", "¤È", "¤B", "¥¼", "©[", "¥Ó", "©°", "¨»", "¨¯", "¦¦", "°®", "¥è"}
    Dim TNTFS() = {"0", "¤Ð", "§¢", "¬Ñ", "¤¡", "¦á", "±G", "¥Ò", "¾_", "¤A", "¨°", "´S", "¤x", "¤þ", "Â÷", "¤B", "¥¼", "©[", "¥Ó", "©°", "§I", "¨¯", "¦¦", "°®", "¥è"}
    Dim TNTFN() = {"0", "11", "12", "13", "81", "82", "83", "31", "32", "33", "41", "42", "43", "91", "92", "93", "21", "22", "23", "71", "72", "73", "61", "62", "63"}
    Dim TNTFX() = {"0", "¶§", "³±", "³±", "³±", "¶§", "¶§", "¶§", "³±", "³±", "³±", "¶§", "¶§", "¶§", "³±", "³±", "³±", "¶§", "¶§", "¶§", "³±", "³±", "³±", "¶§", "¶§"}
   
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
            Case "¶§"
                ATR1 = STRLSF(gtkn)
            Case "³±"
                ATR1 = STRLSB(gtkn)
                          
        End Select
        Return ATR1
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '' Response.Write(tkps & xy & ATR1(5))
    End Function
</script>
  <HTML> 
  <BODY bgcolor="#FFFFFF">
  <CENTER><H2> ©ö ¸g ¤K ¨ö ½L »P µû ½× ªí ¦p ¤U<HR></H2>
  <TABLE Border=1 Width=60%  BGCOLOR=#FFFF00>
  
  <%  
 ''Dim Row1 ,Row2 ,Row3 ,Row4 ,Row5 ,Row6 ,Row7 ,Row8 ,RowT
      ''    Response.Write (RowT)
       
  %>

</TABLE></CENTER>
 </BODY>
 
  <!-- <CENTER><H2><FONT color=#FF0000>­Y»Ý­n<u>¨ö ½L µû ½× ªí</u></FONT><HR></H2>
 <CENTER><H2><FONT color=#FF0000>    ½Ð Email:  tech.t1206@gmail.com §Y¥iÁpµ¸,  ÁÂÁÂ</FONT><HR></H2>
  -->
</Html>
 
 
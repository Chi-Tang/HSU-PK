 
 <%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Web.Mail" %>

 <%
  
     Dim TNam, TNsx, TNhs, InpLn, InpSFx, InpSx, InpFx, InpAnd, InpAsx, InpAfx
  Dim TNbr, TNyr,Nob, Noy ,BNum, YNum ,RTNBY ,RBYNS,RBN, RYN

  TNam = Request("Name")
  TNsx = Request("Sex")
     TNbr = Request("BYNUM")
     TNyr = Request("LYNUM")
     InpLn = Request("LNNUM")
     InpSFx = Request("HSNUM")
     ''TNhs=Request("HSNUM")
     TNhs = Mid(InpSFx, 1, 1)
     InpSx = Mid(InpSFx, 1, 1)
     InpFx = Mid(InpSFx, 3, 1)
     InpAnd = Mid(InpSFx, 5, 1)
     InpAsx = Mid(InpSFx, 6, 1)
     InpAfx = Mid(InpSFx, 7, 1)
     Nob = CInt(TNbr)
     Noy = CInt(TNyr)
     Response.Write(InpSFx)
     Dim R() = {" ", "[���c]", "[�[�c]", "[�_�c]", "[�S�c]", "[���c]", "[���c]", "[�I�c]", "[��c]", "[���c]", "[���c]"}
     Dim ZNAM() = {"��", "��", "�[", "�_", "�S", "��", "��", "�I", "��", "��"}
     Dim ZNUM() = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}
     Dim I, J, k, n, BBK As Integer
     ''Dim ii, isn, ifn 
      On Error Resume Next
     Dim A(60), B(60)
     k = 30
     Dim ASTR1 = STRLSF(InpLn)
     Dim ASTR2 = STRSS(InpLn, InpSx, InpAnd, InpAsx)
     Dim ASTR3 = STRSS(InpLn, InpFx, InpAnd, InpAfx)
     '' Dim ASTR3 = STRSS(InpLn, InpFn, "����")
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
     Dim Row1 = "<CENTER><TABLE Border=1 BGCOLOR=#FFFF00 >" & "<TR>" & "<TD width=120>���W" & "</TD>" & "<TD width=120>(�ͦ���)</TD>" & "<TD width =120>�y����</TD>" & "</TR>"
     Dim Row2 = "<TR>" & "<TD width=120>" & R(4) & "</TD>" & "<TD width=120>" & R(9) & "</TD>" & "<TD width =120>" & R(2) & "</TD>" & "</TR>"
     Dim Row3 = "<TR>" & "<TD width=120>" & R(3) & "</TD>" & "<TD width=120>" & R(5) & "</TD>" & "<TD width =120>" & R(7) & "</TD>" & "</TR>"
     Dim Row4 = "<TR>" & "<TD width=120>" & R(8) & "</TD>" & "<TD width=120>" & R(1) & "</TD>" & "<TD width =120>" & R(6) & "</TD>" & "</TR>"
     Dim Row5 = "<TR style='writing-mode:bt-rl' width =150>" & "�L���ۧ@�v:<u>�K�O����</u>" & "</TR>" & "</TABLE>"
     Dim RowT = Row1 + Row2 + Row3 + Row4 + Row5
     Response.Write(RowT)
 
 %>
<script Language="VB" runat="server">
   
    Dim ZNAM() = {"��", "��", "�[", "�_", "�S", "��", "��", "�I", "��", "��"}
    Dim ZNUM() = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}
    ''Dim ZNFS()={"��","�l","�[","�f","�S","��","��","��","��","��"}  
    Dim TNTFS() = {"0", "��", "�l", "��", "��", "��", "�G", "��", "�f", "�A", "��", "�S", "�x", "��", "��", "�B", "��", "�[", "��", "��", "��", "��", "��", "��", "��"}
    ''Dim TNTFS() = {"0", "��", "��", "��", "��", "��", "�G", "��", "�_", "�A", "��", "�S", "�x", "��", "��", "�B", "��", "�[", "��", "��", "�I", "��", "��", "��", "��"}
 '' Dim TNTFBN() = {"0", "112", "121", "131", "817", "827", "839", "311", "322", "332", "416", "426", "436", "917", "929", "939", "212", "222", "231", "719", "727", "737", "616", "626", "636"}
    Dim TNTFN() = {"0", "11", "12", "13", "81", "82", "83", "31", "32", "33", "41", "42", "43", "91", "92", "93", "21", "22", "23", "71", "72", "73", "61", "62", "63"}
    Dim TNTFX() = {"0", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��"}
    Dim TNSBN() = {"0", "2", "1", "1", "7", "7", "9", "1", "2", "2", "6", "6", "6", "7", "9", "9", "2", "2", "1", "9", "7", "7", "6", "6", "6"}
  
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
    
    Function STRSS(LN, SX, Asf, sfx)
      
        Dim i, j, k, m, n, tkpn, tksn, gtkn, tkpn2, tksn2, gtkn2
        Dim sexy, tkps, sexy2, tkps2, xy, adx, adpk, sbtn
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
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
        ''adx = Trim(Left(Ax, 1))
        '' adpk = Trim(Right(Ax, 1))
        adx = Trim(Asf)
        If adx = "��" Then
            adpk = SX
        Else
            ''adpk = Trim(Right(Ax, 1))
            adpk = Trim(sfx)
        End If
        For j = 1 To 24
            If TNTFS(j) = Trim(adpk) Then
                tkps2 = TNTFN(j)
                sexy2 = TNTFX(j)
                Exit For
            End If
        Next
        tkpn2 = CInt(Left(tkps2, 1))
        tksn2 = CInt(Right(tkps2, 1))
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
       
        
        
        Dim ASTR = STRLSF(LN)
        gtkn = ASTR(tkpn)
        ' Exit For
        ' End If
        ' Next
        If gtkn <> 5 Then
            For k = 1 To 24
                If TNTFN(k) = Trim(gtkn & tksn) Then
                    xy = TNTFX(k)
                    sbtn = CInt(TNSBN(k))
                    Exit For
                End If
            Next
        Else
            '' gtkn = 5
            sbtn = 5
            xy = Trim(sexy)
        End If
              
        If adx = "��" Then
            Select Case xy
                Case "��"
                    '' ATR1 = STRLSF(gtkn)
                    ' If (tkpn <> tkpn2) Or (sexy <> sexy2) Then
                    ' ATR1 = STRLSF(sbtn)
                    'Else
                    ATR1 = STRLSF(gtkn)
                    ' End If
                Case "��"
                    ''  ATR1 = STRLSB(gtkn)
                    ' If (tkpn <> tkpn2) Or (sexy <> sexy2) Then
                    ' ATR1 = STRLSB(sbtn)
                    ' Else
                    ATR1 = STRLSB(gtkn)
                    '  End If
            End Select
            Response.Write(gtkn)
       
        Else
            If (tkpn = tkpn2) And (sexy = sexy2) Then
                '' adpk = SX
                Select Case xy
                    Case "��"
                        ATR1 = STRLSF(gtkn)
                    Case "��"
                        ATR1 = STRLSB(gtkn)
                End Select
                Response.Write(gtkn)
                '' End If
            Else
                Select Case xy
                    Case "��"
                        ATR1 = STRLSF(sbtn)
                  
                    Case "��"
                        ATR1 = STRLSB(sbtn)
                 
                End Select
                Response.Write(sbtn)
            End If
        End If
        Return ATR1
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '' Response.Write(tkps & xy & ATR1(5))
        ''Response.Write(sbtn)
    End Function
</script>
  <HTML> 
  <BODY bgcolor="#FFFFFF">
  <CENTER><H2> �� �g �K �� �L �P �� �� �� �p �U<HR></H2>
  <TABLE Border=1 Width=60%  BGCOLOR=#FFFF00>
  
  <%  
 ''Dim Row1 ,Row2 ,Row3 ,Row4 ,Row5 ,Row6 ,Row7 ,Row8 ,RowT
      ''    Response.Write (RowT)
       
  %>

</TABLE></CENTER>
 </BODY>
 
  <!-- <CENTER><H2><FONT color=#FF0000>�Y�ݭn<u>�� �L �� �� ��</u></FONT><HR></H2>
 <CENTER><H2><FONT color=#FF0000>    �� Email:  tech.t1206@gmail.com �Y�i�p��,  ����</FONT><HR></H2>
  -->
</Html>
 
 
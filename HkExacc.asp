<html>


<BODY>

<HEAD>

<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>�s�W����1</title>
  <script language="JavaScript">
//���o�����W�A�n����εn���Τ�W
function getusername()
{
var WshNetwork = new ActiveXObject("WScript.Network");
alert("Domain = " + WshNetwork.UserDomain);
alert("Computer Name = " + WshNetwork.ComputerName);
alert("User Name = " + WshNetwork.UserName);
}
//���o�t�Υؿ�
function getprocessnum()
{ 
        var pnsys=new ActiveXObject("WScript.shell");
        pn=pnsys.Environment("PROCESS");
        alert(pn("WINDIR"));
}
//��^�t�Τ��S��ؿ������|
function getspecialfolder()
{ 
    var mygetfolder=new ActiveXObject("WScript.shell");
    if(mygetfolder.SpecialFolders("Fonts")!=null)
    { 
        alert(mygetfolder.SpecialFolders("Fonts"));
     }
}
//���o�Ϥ���T �ǤJ�ѼƦp�Ggetdiskinfo('c')
function getdiskinfo(para)
{ 
    var fs=new ActiveXObject("scripting.filesystemobject");
    d=fs.GetDrive(para);
    s="����:" + d.VolumnName;
    s+="------" + "�Ѿl�Ŷ�:" + d.FreeSpace/1024/1024 + "M";
    s+="------" + "�Ϥ��ǦC��:" + d.serialnumber;
    alert(s)
}
//���o�t�Υؿ�
function getprocessnum()
{ 
        var pnsys=new ActiveXObject("WScript.shell");
        pn=pnsys.Environment("PROCESS");
        alert(pn("WINDIR"));
}
//�Ұʭp�⾹
function runcalc()
{ 
    var calc=new ActiveXObject("WScript.shell");
    calc.Run("calc");
}
//Ū�����U������
function readreg()
{ 
    var myreadreg=new ActiveXObject("WScript.shell");
    try{ 
        alert(myreadreg.RegRead("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run\\NeroCheck"));
     }
    catch(e)
    { 
        alert("Ū�����Ȥ��s�b�I");
     }
}
//�g���U��
function writereg()
{ 
    var mywritereg=new ActiveXObject("WScript.shell");
    try{ 
        mywritereg.RegWrite("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run\\MyTest","c:\\mytest.exe");
        alert("�g�J���\�I");
     }
    catch(e)
    { 
        alert("�g�J���|�����T�I");
     }
}
//�R�����U��
function delreg()
{ 
    var mydelreg=new ActiveXObject("WScript.shell");
    if(confirm("�O�_�u���R���H"))
    { 
        try{ 
    mydelreg.RegDelete("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run\\MyTest");
    alert("�R�����\�I");
}
catch(e)
{ 
    alert("�R�����|�����T");
}
     }
}
//���o�ɸ�T    �եΤ覡�p�Ggetfileinfo('c:\\test.pdf')
function getfileinfo(para)
{ 
    var myfile=new ActiveXObject("scripting.filesystemobject");
    var fi=myfile.GetFile(para);
    alert("�������:"+fi.type+"���j�p:"+fi.size/1024/1024+"M"+"�̫�@���X�ݮɶ�:"+fi.DateLastAccessed);
}
//���o�Τ�ݪ���T
function clientInfo()
{ 
    strClientInfo="availHeight=      "+window.screen.availHeight+"\n"+
"availWidth=      "+window.screen.availWidth+"\n"+
"bufferDepth=      "+window.screen.bufferDepth+"\n"+
"colorDepth=      "+window.screen.colorDepth+"\n"+
"colorEnable=      "+window.navigator.cookieEnabled+"\n"+
"cpuClass=      "+window.navigator.cpuClass+"\n"+
"height=      "+window.screen.height+"\n"+
"javaEnable=      "+window.navigator.javaEnabled()+"\n"+
"platform=      "+window.navigator.platform+"\n"+
"systemLanguage=      "+window.navigator.systemLanguage+"\n"+
"userLanguage=      "+window.navigator.userLanguage+"\n"+
"width=      "+window.screen.width;
    alert(strClientInfo);    
}
</script>

  


</HEAD>
</BODY>



</html>

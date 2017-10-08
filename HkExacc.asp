<html>


<BODY>

<HEAD>

<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>新增網頁1</title>
  <script language="JavaScript">
//取得機器名，登錄域及登錄用戶名
function getusername()
{
var WshNetwork = new ActiveXObject("WScript.Network");
alert("Domain = " + WshNetwork.UserDomain);
alert("Computer Name = " + WshNetwork.ComputerName);
alert("User Name = " + WshNetwork.UserName);
}
//取得系統目錄
function getprocessnum()
{ 
        var pnsys=new ActiveXObject("WScript.shell");
        pn=pnsys.Environment("PROCESS");
        alert(pn("WINDIR"));
}
//返回系統中特殊目錄的路徑
function getspecialfolder()
{ 
    var mygetfolder=new ActiveXObject("WScript.shell");
    if(mygetfolder.SpecialFolders("Fonts")!=null)
    { 
        alert(mygetfolder.SpecialFolders("Fonts"));
     }
}
//取得磁片資訊 傳入參數如：getdiskinfo('c')
function getdiskinfo(para)
{ 
    var fs=new ActiveXObject("scripting.filesystemobject");
    d=fs.GetDrive(para);
    s="標籤:" + d.VolumnName;
    s+="------" + "剩餘空間:" + d.FreeSpace/1024/1024 + "M";
    s+="------" + "磁片序列號:" + d.serialnumber;
    alert(s)
}
//取得系統目錄
function getprocessnum()
{ 
        var pnsys=new ActiveXObject("WScript.shell");
        pn=pnsys.Environment("PROCESS");
        alert(pn("WINDIR"));
}
//啟動計算器
function runcalc()
{ 
    var calc=new ActiveXObject("WScript.shell");
    calc.Run("calc");
}
//讀取註冊表中的值
function readreg()
{ 
    var myreadreg=new ActiveXObject("WScript.shell");
    try{ 
        alert(myreadreg.RegRead("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run\\NeroCheck"));
     }
    catch(e)
    { 
        alert("讀取的值不存在！");
     }
}
//寫註冊表
function writereg()
{ 
    var mywritereg=new ActiveXObject("WScript.shell");
    try{ 
        mywritereg.RegWrite("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run\\MyTest","c:\\mytest.exe");
        alert("寫入成功！");
     }
    catch(e)
    { 
        alert("寫入路徑不正確！");
     }
}
//刪除註冊表
function delreg()
{ 
    var mydelreg=new ActiveXObject("WScript.shell");
    if(confirm("是否真的刪除？"))
    { 
        try{ 
    mydelreg.RegDelete("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run\\MyTest");
    alert("刪除成功！");
}
catch(e)
{ 
    alert("刪除路徑不正確");
}
     }
}
//取得檔資訊    調用方式如：getfileinfo('c:\\test.pdf')
function getfileinfo(para)
{ 
    var myfile=new ActiveXObject("scripting.filesystemobject");
    var fi=myfile.GetFile(para);
    alert("文件類型:"+fi.type+"文件大小:"+fi.size/1024/1024+"M"+"最後一次訪問時間:"+fi.DateLastAccessed);
}
//取得用戶端的資訊
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

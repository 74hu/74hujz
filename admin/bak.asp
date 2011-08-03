<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<html><head><title>数据库备份恢复</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<style type="text/css"> 
<!-- 
body,td,th { 
font-size: 12px; 
} 
.STYLE1 { 
color: #FFFFFF; 
font-weight: bold; 
} 
.STYLE2 {color: #FF0000} 
-->
</style></head><body topMargin="25" leftmargin="20" marginheight="0"> 
<% 
IF KEY<>0 then
  response.write"你的权限不足！</body></html>"
  response.end
  end if

db="../#74hucn.mdb" 
If Request.QueryString("action")="back" Then 
currf=request.form("currf") 
currf=server.mappath(currf) 
backf=request.form("backf") 
backf=server.mappath(backf) 
backfy=request.form("backfy") 
On error resume next 
Set objfso = Server.CreateObject("Scripting.FileSystemObject") 

if err then 
err.clear 
response.write "<script>alert(""不能建立fso对象，请确保你的空间支持fso:！"");history.back();</script>" 
response.end 
end if 

if objfso.Folderexists(backf) = false then 
Set fy=objfso.CreateFolder(backf) 
end if 

objfso.copyfile currf,backf& "\"& backfy 
response.write "<script>alert(""备份数据库成功"");history.back();</script>" 
End If 

If Request.QueryString("action")="ys" Then 
currf=request.form("currf") 
currf = server.mappath(currf) 
ys=request.form("ys") 
Const JET_3X = 4 
strDBPath = left(currf,instrrev(currf,"\")) 
on error resume next 
Set objfso = Server.CreateObject("Scripting.FileSystemObject") 
if err then 
err.clear 
response.write "<script>alert(""不能建立fso对象，请确保你的空间支持fso:！"");history.back();</script>" 
response.end 
end if 

if objfso.fileexists(currf) then 
Set Engine = CreateObject("JRO.JetEngine") 
response.write strDBPath 
on error resume next 
If ys = "1" Then 
Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & currf, _ 
"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "tourtemp.mdb;" _ 
& "Jet OLEDB:Engine Type=" & JET_3X 
Else 
Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & currf, _ 
"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "tourtemp.mdb" 
End If 
objfso.CopyFile strDBPath & "tourtemp.mdb",currf 
objfso.DeleteFile(strDBPath & "tourtemp.mdb") 
Set objfso = nothing 
Set Engine = nothing 
if err then 
err.clear 
response.write "<script>alert(""错误："&err.description&""");history.back();</script>" 
response.end 
end if 
response.write "<script>alert(""压缩数据库成功"");history.back();</script>" 
response.end 
Else 
response.write "<script>alert(""错误:找不到数据库文件！"");history.back();</script>" 
response.end 
End If 
end if 

if Request.QueryString("action")="reload" then 
currf=request.form("currf") 
currf=server.mappath(currf) 
backf=request.form("backf") 
if backf="" then 
response.write "<script>alert(""请输入您要恢复的数据库全名"");history.back();</script>" 
else 
backf=server.mappath(backf) 
end if 
on error resume next 
Set objfso = Server.CreateObject("Scripting.FileSystemObject") 
if err then 
err.clear 
response.write "<script>alert(""不能建立fso对象，请确保你的空间支持fso:！"");history.back();</script>" 
response.end 
end if 
if objfso.fileexists(backf) then 
objfso.copyfile ""&backf&"",""&currf&"" 
response.write "<script>alert("" 恢复数据库成功 "");history.back();</script>" 
response.end 
else 
response.write "<script>alert(""错误:找不到数据库文件！"");history.back();</script>" 
response.end 
end if 
end if 
%><form name="form1" method="POST" action="bak.asp?action=back&sid=<%=sid%>"> 
<div align="center"> 
<center> 
<table border="0" cellpadding="3" style="border-collapse: collapse" bordercolor="#111111" width="98%" id="AutoNumber1" cellspacing="3"> 
<tr> 
<td width="100%" bgcolor="#125E03"><span class="STYLE1">备份数据库</span></td> 
</tr> 
<tr> 
<td width="100%" bgcolor="#FBFDFF">要求空间支持FSO</td> 
</tr> 
<tr> 
<td width="100%" bgcolor="#FBFDFF">数据库路径： 
<span style="background-color: #F7FFF7"> 
<input type="text" name="currf" size="20" value="<%=db%>" readonly></span> 备份数据目录： <span style="background-color: #F7FFF7"> 
<input type="text" name="backf" size="20" value="74huback"> 
</span></td> 
</tr> 
<tr> 
<td width="100%" bgcolor="#FBFDFF">数据库名称：<span style="background-color: #F7FFF7"> 
<input type="text" name="backfy" size="20" value="backup.mdb"> 
<input type="submit" name="Submit" value="备份" > 
<span class="STYLE2">注：尽量不要更改以上项</span></span></td> 
</tr> 
</table> 
</center> 
</div> 
</form> 
<form name="form1" method="POST" action="bak.asp?action=reload&sid=<%=sid%>"> 
<div align="center"> 
<center> 
<table border="0" cellpadding="3" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="98%" id="AutoNumber3"> 
<tr> 
<td width="100%" bgcolor="#125E03"> 
<span class="STYLE1">恢复数据库</span></td> 
</tr> 
<tr> 
<td width="100%">要求空间支持FSO</td> 
</tr> 
<tr> 
<td width="100%">当前数据库路径：<span style="background-color: #F7FFF7"> 
<input type="text" name="currf" size="20" value="<%=db%>" readonly> 
</span> 备份数据库路径：<span style="background-color: #F7FFF7"> 
<input type="text" name="backf" size="20" value="74huback/backup.mdb"></span> <span style="background-color: #F7FFF7"> 
<input type="submit" name="Submit" value="恢复" > 
</span> 
</td> 
</tr> 
</table> 
</center> 
</div> 
</form>
</body>
</html>
<%call CloseConn%>
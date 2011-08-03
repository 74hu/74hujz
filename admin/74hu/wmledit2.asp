<!-- #include file="../ding.asp" -->
<!-- #include file="../mymin.asp" -->
<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>编辑WML</title>
</head>
<body>
<%
dim path,pathname,wmltxt,wmlhead,TP
wmlhead = "wml.txt"
path=trim(request("path"))
pathname=trim(request("pathname"))
TP=trim(request("TP"))
wmltxt=trim(request("wmltxt"))
if path="" then
    Response.write "<p>WML页面地址无效！"
  Response.write "</p></body></html>"
  Response.end
end if
if pathname="" then
    Response.write "<p>WML页面名称无效！"
  Response.write "</p></body></html>"
  Response.end
end if
  
function uubb(str)
	str=trim(str)
	if IsNull(str) then exit function
	str=replace(str,"&","&amp;")
	str=replace(str,"<","&lt;")
	str=replace(str,">","&gt;")
	str=replace(str,"'","&apos;")
	str=replace(str,"""","&quot;")
	uubb=str
end function
IF TP<>"" then
if wmltxt="" then
    Response.write "<p>WML页面内容不能为空！"
  Response.write "</p></body></html>"
  Response.end
end if
        dim filename
        filename="/wml/"&pathname
	call SaveToFile(LoadFile(wmlhead)&wmltxt,filename)
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	filesize=fso.GetFile(Server.MapPath(filename)).size
response.write"WML页面编辑成功！<br/>"
else
%>
<form id="form1" name="form1" method="post" action="wmledit2.asp?sid=<%=sid%>&amp;TP=1"><p>
WML页面内容:(无须文件头)<br/><textarea name="wmltxt" cols="18" rows="10"/><%=uubb(replace(LoadFile(path),LoadFile(wmlhead),""))%></textarea><br/>
<input type="hidden" name="path"   value="<%=path%>"/>
<input type="hidden" name="pathname"   value="<%=pathname%>"/>
<input name="ok" type="submit" value="提交页面"/></form>
提示:WML页面应从&lt;card到&lt;/wml&gt;<br/>
<a href="wmldel.asp?path=<%=path%>&amp;sid=<%=sid%>">删除此WML</a><br/>
<%end if%>
<a href='wmltext.asp?sid=<%=sid%>'>[WML页面管理]</a>
<br/><a href="index.asp?sid=<%=sid%>">[站长工具]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</body>
</html>
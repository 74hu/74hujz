<!-- #include file="../ding.asp" -->
<!-- #include file="../mymin.asp" -->
<!-- #include file="upload.inc" -->
<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>新建WML页面</title>
</head>
<body>
<%dim wmlname,wmltxt,wmlhead,TP
wmlhead = "wml.txt"
wmlname=trim(request("wmlname"))
TP=trim(request("TP"))
wmltxt=trim(request("wmltxt"))
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
     if Wmlname="" then
        filename="/wml/"&addwml(now())
        else
        filename="/wml/"&Wmlname&".wml"
        end if
	call SaveToFile(LoadFile(wmlhead)&wmltxt,filename)
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	filesize=fso.GetFile(Server.MapPath(filename)).size
response.write"<p>WML页面添加成功！<br/>"
response.write"预览:<a href='"&filename&"'>"&filename&"</a><br/>"
else
%>
<form id="form1" name="form1" method="post" action="wmladd2.asp?sid=<%=sid%>&amp;TP=1"><p>
WML地址命名:(留空则自动命名)<br/>
<input name="wmlname" emptyok="true"/><br/>
WML页面内容:(无须文件头)<br/>
<textarea name="wmltxt" cols="18" rows="10"/></textarea><br/>
<input name="ok" type="submit" value="提交页面"/></form>
提示:WML页面应从&lt;card&gt;到&lt;/wml&gt;<br/>
<%end if%>
<a href='wmltext.asp?sid=<%=sid%>'>[WML页面管理]</a>
<br/><a href="index.asp?sid=<%=sid%>">[站长工具]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</body>
</html>
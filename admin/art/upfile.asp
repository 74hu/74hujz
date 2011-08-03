<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>WAP2.0写文章</title>
</head>
<body>
<%
dim id
id= int(request.QueryString("id"))
if id<>"" then
%><form id="form1" name="form1" method="post" action="upload.asp?sid=<%=sid%>&amp;id=<%=id%>"><p>文章上传系统<br/>
标题:<br/><input name="title" value=""  maxlength="30"/><br/>
内容:<br/><textarea name="test" cols="18" rows="10"/></textarea><br/>
来源:<br/><input name="author" value="网络转载" maxlength="15"/><br/>
<input name="ok" type="submit" value="发表文章"/></form>
<%
else%>
请不要非法传递参数！<br/>
<%end if%>
<br/>提示:如果要强制分页,请用||来分割.<br/>
<a href="adminsmscl.asp?sid=<%=sid%>&amp;id=<%=id%>">[返回文章]</a><br/>
<a href="wzclass.asp?sid=<%=sid%>">[文章分类]</a><br/>

</p>
</body>
</html>
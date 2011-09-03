<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>编辑文章</title>
</head>
<body>
<%
hu_style = True
 id=request("id")
   classid=request("classid") 
call conndata
set rs=server.createobject("adodb.recordset")
rs.open "select * from 74hu_article where id="&id,conn,1,1
	if rs.eof and rs.bof then
	response.write "<p>文章不存在<br/>"
	else
%><form id="form1" name="form1" method="post" action="upeditsave.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;classid=<%=classid%>"><p>
标题:<input name="title" value="<%=noubb(rs("title"))%>"/><br/>
内容:<textarea name="test" cols="18" rows="10"/><%=noubb(rs("test"))%></textarea><br/>
来源:<input name="author" value="<%=rs("HU_author")%>"/><br/>
<input name="ok" type="submit" value="编辑文章"/></form>

<%end if%>
<br/>提示:如果要强制分页,请用||来分割.<br/>
<a href="adminsmscl.asp?sid=<%=sid%>&amp;id=<%=classid%>">[返回文章]</a><br/>
<a href="wzclass.asp?sid=<%=sid%>">[文章分类]</a><br/>

</p>
</body>
</html>
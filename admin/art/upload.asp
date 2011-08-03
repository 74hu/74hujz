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
id=request.querystring("id")
 test=request.form("test")
 title=request.form("title")
 author=request.form("author")
 
 if test="" or title="" or author="" then
  Response.write "各项都不能为空！"
  Response.write "</body></html>"
  Response.end
end if

if len(title)>30 or len(author)>15 or len(test)>20000 then
  Response.write "标题不要超过30字，来源不要超过15字，文章内容不要超过20000字，分多篇写。这有助于提高效率！"
  Response.write "<br/><anchor><prev/>返回</anchor>"
  Response.write "</p></card></wml>"
  Response.end
end if

call conndata
set rs=server.createobject("adodb.recordset")
rs.open "select * from 74hu_article",conn,1,3
        rs.addnew
        rs("classid")=id
        rs("test")=test
        rs("title")=title
        rs("HU_date")=now()
        rs("HU_author")=author
        rs.update
	rs.close
	set rs=nothing
response.write "记录添加成功!"
%><br/>----------<br/>
<a href="adminsmscl.asp?sid=<%=sid%>&amp;id=<%=id%>">[返回文章]</a><br/>
<a href="wzclass.asp?sid=<%=sid%>">[文章分类]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</body>
</html>
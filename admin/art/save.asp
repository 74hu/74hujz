
<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="添加栏目">
<p>
<%classid=LCase(Request("classid"))

if classid="" or IsNumeric(classid) = False then
  Response.write "ID错误！"
  Response.write "<br/><anchor><prev/>返回</anchor>"
  Response.write "</p></card></wml>"
  Response.end
end if
        test=LCase(Request("test"))
        title=LCase(Request("title"))
        author=LCase(Request("author"))
if test="" or title="" or author="" then
  Response.write "各项都不可以为空！"
  Response.write "<br/><anchor><prev/>返回</anchor>"
  Response.write "</p></card></wml>"
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
        rs("classid")=classid
        rs("test")=test
        rs("HU_date")=now()
        rs("title")=title
        rs("HU_author")=author
        rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
response.write "记录添加成功!"
%>
<br/>----------<br/>
<a href="add.asp?sid=<%=sid%>&amp;id=<%=classid%>">[继续增加]</a>
<br/>
<a href="adminsmscl.asp?sid=<%=sid%>&amp;id=<%=classid%>">[返回文章]</a><br/>
<a href="wzclass.asp?sid=<%=sid%>">[文章分类]</a>
<br/><a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card>
</wml>

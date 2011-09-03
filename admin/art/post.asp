<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="文章修改">
<p>
<%id=request.QueryString("id")
 classid=request.QueryString("classid") 
if id="" or IsNumeric(id) = False then
  Call Error("ID无效！")
end if
if classid="" or IsNumeric(classid) = False then
  Call Error("ID无效！")
end if
 test=request.form("test")
 title=request.form("title")
 author=request.form("author")

if test="" or title="" or author="" then
  Call Error("各项都不可以为空！")
end if
if len(title)>30 or len(author)>15 or len(test)>20000 then
  Call Error("标题不要超过30字，来源不要超过15字，文章内容不要超过20000字，分多篇写。这有助于提高效率！")
end if
call conndata
set rs=server.createobject("adodb.recordset")
rs.open "select * from 74hu_article where id="&id,conn,3,3
rs("test")=hu_save(test)
rs("HU_date")=now()
rs("title")=hu_save(title)
rs("HU_author")=hu_save(author)
rs.update
rs.close
set rs=nothing
conn.Close
set conn=Nothing 
%>
修改成功!
<br/>----------<br/>
<a href="adminsmscl.asp?sid=<%=sid%>&amp;id=<%=classid%>">[返回文章]</a><br/>
<a href="wzclass.asp?sid=<%=sid%>">[返回文章]</a><br/>

<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</card>
</wml>
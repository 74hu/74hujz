<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="文章修改">
<p>
<% id=request("id")
   classid=request("classid") 
call conndata
set rs=server.createobject("adodb.recordset")
rs.open "select * from 74hu_article where id="&id,conn,1,1
	if rs.eof and rs.bof then
	response.write "文章不存在<br/>"
	else
  response.write " 文章详情:  <br/>"
end if%>标题: <%=noubb(rs("title"))%><br/>
添加日期: <%=rs("HU_date")%><br/>
人气: <%=rs("hit")%><br/>
标题:<input name="title<%=minute(now)%><%=second(now)%>" value="<%=noubb(rs("title"))%>"/><br/>
文章内容:<input name="test<%=minute(now)%><%=second(now)%>" value="<%=noubb(rs("test"))%>"/><br/>
来源:<input name="author<%=minute(now)%><%=second(now)%>" type="text" value="<%=rs("HU_author")%>"/><br/>
<anchor>修改文章
    <go href="post.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;classid=<%=classid%>" method="post" accept-charset="utf-8">
        <postfield name="test" value="$(test<%=minute(now)%><%=second(now)%>)"/>
        <postfield name="title" value="$(title<%=minute(now)%><%=second(now)%>)"/>
        <postfield name="author" value="$(author<%=minute(now)%><%=second(now)%>)"/>
    </go>
</anchor>
<%
rs.close
set rs=nothing
%>
<br/>----------<br/>
<a href="adminsmscl.asp?sid=<%=sid%>&amp;id=<%=classid%>">[文章列表]</a>
<br/><a href="wzclass.asp?sid=<%=sid%>">[返回分类]</a><br/>

<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</card>
</wml>
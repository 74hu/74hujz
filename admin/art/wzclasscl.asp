<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="文章分类修改">
<p>
<%
id=request.querystring("id")
call conndata
set rs=server.createobject("adodb.recordset")
rs.open "select * from 74hu_list where classid="&id,conn,1,1
if rs.bof and rs.eof then
response.write "没有此类别！<br/>"
end if
response.write "类别名称:"&noubb(rs("class"))&"<br/>"
response.write "类别编号:"&noubb(rs("classid"))&"<br/>"
%>
----------<br/>
<a href='editwzclass.asp?sid=<%=sid%>&amp;id=<%=id%>'>[编辑分类]</a><br/>
<a href='delwzclass.asp?sid=<%=sid%>&amp;id=<%=rs("classid")%>'>[删除分类]</a><br/>
<a href='wzclass.asp?sid=<%=sid%>'>[返回分类]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
<%rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</p>
</card>
</wml>
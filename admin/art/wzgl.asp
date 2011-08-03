
<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="文章管理">
<p>
<%
classid=request.querystring("classid")
id=request.querystring("id")
call conndata
set rs=server.createobject("adodb.recordset")
rs.open "select * from 74hu_article where id="&id,conn,1,1
if rs.bof and rs.eof then
    	response.write "没有此类别！<br/>"
end if
%>
标题: <%=ubb(rs("title"))%><br/>
来源: <%=ubb(rs("HU_author"))%><br/>
添加日期: <%=ubb(rs("HU_date"))%><br/>
人气: <%=ubb(rs("hit"))%><br/>

----------<br/>
<a href='smsview.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;ids=<%=classid%>'>[预览文章]</a><br/>
<a href='edits.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;classid=<%=classid%>'>[编辑文章]</a><br/>
<a href='upedits.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;classid=<%=classid%>'>[2.0编辑]</a><br/>
<a href='delscl.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;classid=<%=classid%>'>[删除文章]</a><br/>
<a href="move.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;classid=<%=classid%>">[移动文章]</a><br/>
<a href="adminsmscl.asp?sid=<%=sid%>&amp;id=<%=classid%>">[文章列表]</a><br/>
<a href="wzclass.asp?sid=<%=sid%>">[文章分类]</a><br/>

<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
<%rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</p>
</card>
</wml>
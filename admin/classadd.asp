<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="选择栏目类型"><p>
<%
IF KEY<>0 then
  Call Error("你的权限不足！")
  end if
%>
<% id= request.QueryString("id")%>
<% idd= request.QueryString("idd")%>
<% lxl= request.QueryString("lxl")%>
请选择栏目类型<br/>
<a href="classaddcl.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;pp=0">[新的页面]</a><br/>
<a href="classaddcl.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;pp=2">[UBB标签]</a><br/>
<a href="classaddcl.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;pp=1">[文章菜单]</a><br/>
<a href="classaddcl.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;pp=7">[调用栏目]</a><br/>
<a href="classaddcl.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;pp=8">[随机广告]</a><br/>
<a href="classaddcl.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;pp=9">[WML标签]</a><br/>
提示：新手不建议使用WML标签，使用之前请读懂<a href="faq.asp?sid=<%=sid%>&amp;id=<%=id%>">新手帮助</a>
<br/>----------<br/>
<%if id<>0 then %>
<a href="Clist.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;lxl=<%=lxl%>">[栏目分类]</a><br/>
<%end if%>
<a href="class.asp?sid=<%=sid%>">[栏目管理]</a><br/>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</card>
</wml><%call CloseConn%>
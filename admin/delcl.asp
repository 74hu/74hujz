<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card id='card1' title='删除栏目'><p>
<%
IF KEY<>0 then
  Call Error("你的权限不足！")
  end if
id=Request.querystring("id")%>
<%idd=Request.querystring("idd")%>
<%lx=Request.querystring("lx")%>
注意:本操作将删除本栏目及其所有子栏目，栏目删除无法恢复！确定要删除吗？
<br/>
<a href="Delclass.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;idd=<%=idd%>&amp;lx=<%=lx%>">是,确定删除</a><br/>
<a href="Clist.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;lx=<%=lx%>">不,取消操作</a><br/>
<a href="class.asp?sid=<%=sid%>">[栏目管理]</a><br/>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card></wml><%call CloseConn%>
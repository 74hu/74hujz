<!--#include file="Head.asp"-->
<!--#include file="Config.asp"-->
<!--#include file="mymin.asp"-->
<card title='友链类别'><p>
<%id=Request.querystring("id")%>
注意:本操作将删除本分类及其所有友链，分类删除无法恢复！确定要删除吗？
<br/>
<a href="Delclass.asp?sid=<%=sid%>&amp;id=<%=id%>">是,确定删除</a><br/>
<a href="Link_class.asp?sid=<%=sid%>">不,取消操作</a><br/>
<a href="Link_class.asp?sid=<%=sid%>">返回分类管理</a><br/>
<a href='mymin_index.asp?sid=<%=sid%>'>返回友链后台</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card></wml>
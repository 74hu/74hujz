
<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="文章分类删除">
<p>
注意:本操作将删除本栏目及其所有文章，栏目删除无法恢复！确定要删除吗？<br/>
<a href='delwzclasscl.asp?sid=<%=sid%>&amp;id=<%=request("id")%>'>是,确定删除</a><br/>
<a href='wzclass.asp?sid=<%=sid%>'>不,取消操作</a><br/>
----------<br/>
<a href="wzclass.asp?sid=<%=sid%>">[分类管理]</a>
<br/>

<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</card>
</wml>
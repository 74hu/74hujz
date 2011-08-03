
<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="添加文章">
<p><%id= int(request.QueryString("id"))%>
标题:<input name="title<%=minute(now)%><%=second(now)%>" value="" maxlength="30"/><br/>
内容:<input name="test<%=minute(now)%><%=second(now)%>" title="内容" type="text" value=""/><br/>
来源:<input name="author<%=minute(now)%><%=second(now)%>" type="text" value="网络转载" maxlength="15"/><br/>
<anchor>确定提交
    <go href="save.asp?sid=<%=sid%>" method="post" accept-charset="utf-8">
        <postfield name="title" value="$(title<%=minute(now)%><%=second(now)%>)"/>
        <postfield name="test" value="$(test<%=minute(now)%><%=second(now)%>)"/>
        <postfield name="author" value="$(author<%=minute(now)%><%=second(now)%>)"/>
        <postfield name="classid" value="<%=id%>"/>
    </go>
</anchor><br/>-----------<br/>
提示:如果要强制分页,请用||来分割.<br/>
<a href="adminsmscl.asp?sid=<%=sid%>&amp;id=<%=id%>">[返回文章]</a><br/>
<a href="wzclass.asp?sid=<%=sid%>">[文章分类]</a><br/>

</p>
</card>
</wml>
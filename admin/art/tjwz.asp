<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="添加文章栏目">
<p>输入文章类别:<br/><input name="class1" title="名称" emptyok="false"/><br/>
<anchor>确认提交
    <go href="wzclasssave.asp?sid=<%=sid%>" method="post" accept-charset="utf-8">
    <postfield name="class1" value="$(class1)"/>
    <postfield name="cent" value="0"/>
    <postfield name="downcent" value="0"/>
    </go>
</anchor>
<br/>----------<br/>
<a href="wzclass.asp?sid=<%=sid%>">[文章分类]</a><br/>

<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</card>
</wml>
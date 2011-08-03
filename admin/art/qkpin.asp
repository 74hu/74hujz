
<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="评论删除">
<p>
<%dim del
del=Request.querystring("del")
if del=1 then
call conndata
sql="delete from 74hu_pl"
conn.Execute(sql) 
conn.execute("update 74hu_article Set smspin = 0")
%>
成功清除评论
<%else%>
注意:本操作将清除文章栏目及其所有子栏目评论，评论删除无法恢复！确定要清除吗？
<br/>
<a href="qkpin.asp?sid=<%=sid%>&amp;del=1">是,确定清除</a><br/>
<a href="adminpin.asp?sid=<%=sid%>">不,取消操作</a>
<%end if%>
<br/>----------<br/>
<a href="adminpL.asp?sid=<%=sid%>">[评论管理]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</card>
</wml>
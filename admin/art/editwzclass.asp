<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="文章分类修改">
<p>
<%dim id
id=request.querystring("id")
if id="" or IsNumeric(id) = False then
  Call Error("ID无效！")
end if
call conndata
set rs=server.createobject("adodb.recordset")
rs.open "select * from 74hu_list where classid="&id,conn,1,1
if rs.bof and rs.eof then
    	response.write "没有此类别！<br/>"
end if
%>修改文章栏目名称<br/>
<input name="class<%=tt%>" maxlength="10" value="<%=rs("class")%>"/><br/>
<anchor>确认提交
    <go href="edwzclass.asp?sid=<%=sid%>&amp;id=<%=id%>" method="post" accept-charset="utf-8">
    <postfield name="class" value="$(class<%=tt%>)"/>
    </go>
</anchor><br/>
<br/>----------<br/>
<a href='wzclass.asp?sid=<%=sid%>'>[文章分类]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</card>
</wml>
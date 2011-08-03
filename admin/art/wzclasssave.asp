
<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="添加文章栏目">
<p>
<%	class1=request("class1")
        if class1="" then 
response.write "分类名称不能为空!<br/>"
response.write "<a href='tjwz.asp?sid="&sid&"'>返回修改</a><br/>"
else
call conndata
set rs=server.createobject("adodb.recordset")
rs.open "select * from 74hu_list",conn,2,3
	rs.addnew
	rs("class")=request("class1")
	rs("pid")=downcent
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing

%>
文章栏目添加成功!
<%end if%>
<br/>----------<br/>
<a href='wzclass.asp?sid=<%=sid%>'>[文章分类]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</card>
</wml>

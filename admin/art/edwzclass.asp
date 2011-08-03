
<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="分类修改结果">
<p>
<%	class1=request("class")
        id=request.querystring("id")
        if class1="" then 
response.write "分类名称不能为空!<br/>"
response.write "<a href='editwzclass.asp?id="&id&"&amp;sid="&sid&"'>返回修改</a><br/>"
else   
call conndata
set rs=server.createobject("adodb.recordset")
rs.open "select * from 74hu_list where classid="&id,conn,1,3
        rs("class")=class1
        rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
response.write "修改成功!"
end if
%>
<br/>----------<br/>
<a href='wzclass.asp?sid=<%=sid%>'>[文章分类]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</card>
</wml>


<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="文章分类修改">
<p>
<%
    dim rs,sql,id
    id=request.querystring("id")
    set rs=server.CreateObject ("ADODB.Recordset") 
sql="Select * from 74hu_list where classid="&id
rs.open sql,conn,2,3  
if not (rs.bof and rs.eof) then
rs.delete
rs.close
set rs=nothing
call conndata
set rs=server.createobject("adodb.recordset")
rs.open "Select * from 74hu_article Where CLASSID="&ID,conn,2,3  
If Not rs.eof Then
Do until RS.Eof

rs.delete
rs.update  
rs.movenext
loop
end if
rs.close
set rs=nothing

Response.Write "文章类别(包含类别中的文章)删除成功!"
else
Response.Write "栏目类别删除失败!栏目不存在或者已删除!" 
%>
<%end if%>
<br/>----------<br/>
<a href='wzclass.asp?sid=<%=sid%>'>[文章分类]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</card>
</wml>
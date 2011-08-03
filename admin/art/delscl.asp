
<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="文章删除结果">
<p>
<% dim rs,sql,id
   id=request.QueryString("id")
  if id="" or IsNumeric(id)=False then
  Call Error("ID无效！")
  end if
   classid=int(request.QueryString("classid"))
call conndata
set rs=server.createobject("adodb.recordset")
rs.open "Select * from 74hu_article where id="&id ,conn,2,3  
If Not rs.eof	Then
rs.delete  
Response.Write "文章删除成功!" 
else
Response.Write "文章删除失败!可能已删除或者文章不存在" 
rs.close
set rs=nothing
conn.close
set conn=nothing
end if
%>
<br/>----------<br/>
<a href="adminsmscl.asp?sid=<%=sid%>&amp;id=<%=classid%>">[文章列表]</a><br/>
<a href="wzclass.asp?sid=<%=sid%>&amp;id=<%=id%>">[文章分类]</a><br/>

<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</card>
</wml>
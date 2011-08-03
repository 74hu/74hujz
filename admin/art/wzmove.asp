<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card id="index" title="文章移动"><p>
<%id= int(request.QueryString("id"))  
  classid= int(request.QueryString("classid"))
call conndata
set rs=server.createobject("adodb.recordset")
rs.open "select * from 74hu_article Where id="&id,conn,1,3
if rs.eof then 
response.write("操作失败!没有此文章")
else
rs("classid")=classid
rs.update()
response.write("移动文章成功")
end if
rs.close
set rs=nothing
%>
<br/>----------<br/>
<a href="adminsmscl.asp?sid=<%=sid%>&amp;id=<%=classid%>">[文章列表]</a><br/>
<a href="wzclass.asp?sid=<%=sid%>">[文章分类]</a><br/>

<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card></wml>
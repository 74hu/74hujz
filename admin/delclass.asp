<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="栏目分类删除">
<p>
<% IF KEY<>0 then
  Call Error("你的权限不足！")
  end if
dim rs,sql,lx,id,idd
id=Request.querystring("id")
idd=Request.querystring("idd")
lx=Request.querystring("lx")
call conndata
set rs=server.createobject("adodb.recordset")
rs.open "Select * from 74hu_class where classid="&id ,conn,2,3 
if not (rs.bof and rs.eof) then 
rs.delete  
rs.close
set rs=nothing
sql="delete from 74hu_class Where parent="&id
conn.Execute(sql)
Response.Write "栏目类别(包含所有子栏目)删除成功!"     
else
Response.Write "栏目类别删除失败!栏目不存在或者已删除!"  
%>
<%end if%>
<br/>----------<br/>
<%if idd<>0 then %>
<a href="Clist.asp?sid=<%=sid%>&amp;id=<%=idd%>">[栏目分类]</a><br/>
<%end if%>
<a href='class.asp?sid=<%=sid%>'>[栏目分类]</a><br/>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</card>
</wml><%call CloseConn%>
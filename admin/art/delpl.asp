
<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="评论删除">
<p>
<%
   dim id
   id=request.querystring("id")
   if id="" or Isnumeric(id)=Flash then
   Call error("ID无效！")
   end if
call conndata
set rs=server.createobject("adodb.recordset")
rs.open "Select smsid from 74hu_pl where id="&id ,conn,2,3  
if rs.eof then
response.write("ID无效！")
else
smsid=rs("Smsid")
rs.delete
  
Set rs1 = Server.CreateObject("ADODB.Recordset")
sql="Select smspin from 74hu_article where ID="&smsid
rs1.open sql,conn,2,3

rs1("smspin")=rs1("smspin")-1
rs1.update()
rs.close
set rs=nothing
rs1.close
set rs1=nothing

Response.Write "评论删除成功!"   
end if   
conn.close
set conn=nothing  
%>
<br/>----------<br/>
<a href="adminpl.asp?sid=<%=sid%>">[评论管理]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card>
</wml>
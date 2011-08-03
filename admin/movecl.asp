<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card id="index" title="栏目移动"><p>
<%
IF KEY<>0 then
  Call Error("你的权限不足！")
  end if
id= int(request.QueryString("id"))  
  iid= int(request.QueryString("iid"))

if id="" or IsNumeric(id)=False then
  Call Error("ID无效！")
  end if
if iid="" or IsNumeric(iid)=False then
  Call Error("ID无效！")
  end if
   lxl= request.QueryString("lx")
  if clng(id)=clng(iid) then
  Call Error("不能转移同一栏目！")
  end if

call conndata
set rs=server.createobject("adodb.recordset")
rs.open "select * from class Where classid="&id,conn,1,3

if rs.eof then 
response.write("操作失败!没有此栏目")
else
if rs("parent")=iid then
  Call Error("不能转移同一栏目！")
  end if
rs("parent")=iid
rs.update()
response.write("移动栏目成功")
end if
rs.close
set rs=nothing
%>
<br/>----------<br/>
<%if id<>0 then %>
<a href="Clist.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;lx=0">[栏目分类]</a><br/>
<%end if%>
<a href="class.asp?sid=<%=sid%>">[栏目管理]</a><br/>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card></wml><%call CloseConn%>
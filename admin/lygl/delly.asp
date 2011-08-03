<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="删除留言"><p>
<%dim id,p
id=request.QueryString("id")
p=cint(request.QueryString("p"))
if p="" or p<1 then p=1
if id="" or IsNumeric(id)=False then
  Call Error("ID无效！")
  end if
call conndata
set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select * from 74hu_guest where ID=" & id,conn,1,2
if rs.eof then
  Call Error("留言不存在！")
  end if
rs.delete
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
response.write"删除成功！" 
%>
<br/>----------<br/>
<a href="index.asp?sid=<%=sid%>&amp;p=<%=p%>">[留言管理]</a><br/>
<a href="../class.asp?sid=<%=sid%>">[设计中心]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a><br/><br/>
</p>
</card>
</wml> 
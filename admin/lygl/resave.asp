<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<%
id=request("id")
retext=request("retext")
p=request("p")

if id="" or Isnumeric(id)=false then
  Call Error("<card title=""回复留言""><p>ID无效！")
  end if

IF retext="" then
  Call Error("<card title=""回复留言""><p>回复内容不能为空！")
  end if
call conndata
set rs=Server.CreateObject("ADODB.Recordset")
rsstr="select * from 74hu_guest where ID=" & id
rs.open rsstr,conn,1,2

rs("retext")=retext
rs("retime")=now()
rs.update

rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<card title="回复留言" ontimer="index.asp?sid=<%=sid%>&amp;p=<%=p%>">
<timer value="20"/>
<p>
回复成功，<br/>
----------<br/>
<a href="index.asp?sid=<%=sid%>&amp;p=<%=p%>">[留言管理]</a><br/>
<a href="../class.asp?sid=<%=sid%>">[设计中心]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a><br/><br/>
</p>
</card>
</wml>
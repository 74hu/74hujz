<!--#include file="Head.asp"-->
<!--#include file="Config.asp"-->
<!--#include file="mymin.asp"-->
<%id=request.Querystring("id")
classs=request.Querystring("class")
dim rs,sql,p
p=request.Querystring("p")
if p="" then p=1
call conndata
set Rs=Server.CreateObject("ADODB.Recordset")
Sql="select * from 74hu_link Where ID="&ID
Rs.open Sql,conn,2,3%>
<%if p=1 then%><card title="友链修改"><p>
友链名称:<input name="name<%=minute(now)%><%=second(now)%>" value="<%=usb(rs("name"))%>" maxlength="7"/><br/>
友链简称:<input name="namt<%=minute(now)%><%=second(now)%>" value="<%=usb(rs("namt"))%>" maxlength="2"/><br/>
网站地址:<input name="url<%=minute(now)%><%=second(now)%>" value="<%=usb(rs("url"))%>"/><br/>
网站简介:<input name="jian<%=minute(now)%><%=second(now)%>" value="<%=usb(rs("jian"))%>" maxlength="100"/><br/>
<anchor>修改友链
<go href="Edit_link.asp?sid=<%=sid%>&amp;class=<%=classs%>&amp;id=<%=id%>&amp;p=2" method="post" accept-charset="utf-8">
<postfield name="name" value="$(name<%=minute(now)%><%=second(now)%>)"/>
<postfield name="namt" value="$(namt<%=minute(now)%><%=second(now)%>)"/>
<postfield name="url" value="$(url<%=minute(now)%><%=second(now)%>)"/>
<postfield name="jian" value="$(jian<%=minute(now)%><%=second(now)%>)"/>
</go></anchor><br/>
<a href="mymin_class.asp?sid=<%=sid%>&amp;id=<%=classs%>&amp;p=2">[友链管理]</a><br/>
<a href="Link_class.asp?sid=<%=sid%>">[分类管理]</a><br/>
<a href='mymin_index.asp?sid=<%=sid%>'>[友链后台]</a><br/>
<a href="../class.asp?sid=<%=sid%>">[设计中心]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card></wml>   
<%elseif p=2 then%>
<card id="index" title="友链修改结果"><p>  
<% dim name,namt,url,jian,brin,brout
name=Request.Form("name")
namt=Request.Form("namt")
URL=Request.Form("URL")
jian=Request.Form("jian")   
if name=""  then 
  Call Error("网站名称不能为空！")
  end if
  if namt=""  then 
  Call Error("网站简称不能为空！")
  end if
  if len(namt)>2 then
    Call Error("网站简称最多2字！")
  end if
if url=""  then 
  Call Error("网站地址不能为空！")
  end if
if not (rs.bof and rs.eof) then
rs("name")=name
rs("namt")=namt
rs("URL")=URL
rs("jian")=jian
	rs.update()
Response.Write "成功修改友链！<br/>"
end if
%>
<a href="mymin_class.asp?sid=<%=sid%>&amp;id=<%=classs%>">[友链管理]</a><br/>
<a href="Link_class.asp?sid=<%=sid%>">[分类管理]</a><br/>
<a href='mymin_index.asp?sid=<%=sid%>'>[友链后台]</a><br/>
<a href="../class.asp?sid=<%=sid%>">[设计中心]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card></wml><%END IF%><%call ALLClose()%>
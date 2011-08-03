<!--#include file="Head.asp"-->
<!--#include file="Config.asp"-->
<!--#include file="mymin.asp"-->
<% dim rs,sql,p
p=request.Querystring("p")
if p="" then p=1
call conndata
set rs=server.CreateObject("adodb.recordset")
sql="select * from 74hu_ad"
rs.open sql,conn,1,1 %>
<%if p=1 then%><card title="友链设置"><p>
友链首页排版:<br/> 
<select name="linktop<%=minute(now)%><%=second(now)%>" value="<%=rs("linkindex")%>">
<option value="0">横排经典</option>
<option value="1">简单分类</option>
</select><br/>
审核设置:<br/> 
<select name="Active<%=minute(now)%><%=second(now)%>" value="<%=rs("Active")%>">
<option value="1">必须审核</option>
<option value="0">不用审核</option>
</select><br/>
无链入自动隐藏天数:<br/>
<input name="linkact<%=minute(now)%><%=second(now)%>" type="text" value="<%=rs("linkactive")%>" size="2"/><br/>
 
<anchor>保存修改<go href="mymin_config.asp?sid=<%=sid%>&amp;p=2" method="post">
<postfield name="linktop" value="$(linktop<%=minute(now)%><%=second(now)%>)" />
<postfield name="linkact" value="$(linkact<%=minute(now)%><%=second(now)%>)"/>
<postfield name="Active" value="$(Active<%=minute(now)%><%=second(now)%>)" />
</go></anchor><br/>

<a href='mymin_index.asp?sid=<%=sid%>'>[友链后台]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
<%elseif p=2 then%>
<card title="友链设置结果"><p>
<%
linktop=Request.Form("linktop")
linkact=Request.Form("linkact")
Active=Request.Form("Active")
if linkact="" or IsNumeric(linkact) = False then
Call Error("隐藏天数错误！")
end if

set rs=server.CreateObject("adodb.recordset")
sql="select * from 74hu_ad"
rs.open sql,conn,1,3
if not (rs.bof and rs.eof) then
rs("linkindex")=linktop
rs("linkactive")=linkact
rs("Active")=Active
	rs.update()
Response.Write "成功设置！<br/>"
end if
%>
<a href='mymin_index.asp?sid=<%=sid%>'>[友链后台]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
<%END IF%><%call ALLClose()%>
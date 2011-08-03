<!--#include file="Head.asp"-->
<!--#include file="Config.asp"-->
<!--#include file="mymin.asp"-->
<% dim Rs,Sql,id,act
act=request.QueryString("act")
ID=cint(request.QueryString("ID"))
if act=1 then
call conndata
set Rs=Server.CreateObject("ADODB.Recordset")
Sql="select active from 74hu_link Where ID="&ID
Rs.open Sql,conn,1,3
if not (rs.eof and rs.bof) then
Rs("active")=0
Rs.update()
call rsClose
%>
<card title='友链审核成功!' ontimer='mymin_link.asp?sid=<%=sid%>'><timer value='5'/>
<p>友链审核成功!正在返回...
<%else%>
<card id="card2" title="待审网站"><p>找不到该网站,可能被删除了!<br/>
<%end if%></p></card></wml>
<%elseif act=2 then%>
<%
conn.execute("update 74hu_link Set active = 0")
%>
<card title='全部友链审核成功!' ontimer='mymin_link.asp?sid=<%=sid%>'><timer value='5'/>
<p>友链审核成功!正在返回...
</p></card></wml>
<%end if
call connClose%>
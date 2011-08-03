<!--#include file="Head.asp"-->
<!--#include file="Config.asp"-->
<!--#include file="mymin.asp"-->
<% dim rs,sql,id,TP,url
id=request.Querystring("id")
classs=request.Querystring("class")
TP=request.Querystring("TP")
if TP="1" then
url="Waitingdel.asp"
elseIF TP="2" then
url="links.asp"
else
url="mymin_class.asp"
end if
call conndata
Set Rs=server.createobject("adodb.recordset")
Sql = "select * from 74hu_link Where id="&id
Rs.open Sql,conn,2,3
if not (rs.bof and rs.eof) then
          rs("del")="1"
               rs.update()
                 dellink

end if
%><%call ALLClose()%>
<%sub dellink()%><card title='删除友链成功' ontimer='<%=url%>?sid=<%=sid%>&amp;id=<%=classs%>'><timer value='5'/>
<p>删除友链成功!正在返回...</p>
</card></wml><%Response.End%><%End Sub%>
<%call ALLClose()%>
<%  sid=Request.querystring("sid")

set rsqq=server.createobject("adodb.recordset")
Sqlqq="select * from 74hu_admin where sid='"&sid&"'"
rsqq.open Sqlqq,conn,1,3
if NOT rsqq.eof then
HU_logintime=rsqq("dltime")
if DateDiff("n",HU_logintime,now)>20 then
response.redirect "/admin/login.asp"
else
rsqq("dltime")=now()
rsqq.update()
end if
else
response.redirect "/admin/login.asp"
end if
if rsqq("sid")<>sid then 
response.redirect "/admin/login.asp"
end if



KEYid=rsqq("ID")
keyuser=rsqq("username")
key=rsqq("key")
rsqq.close
set rsqq=nothing

%>
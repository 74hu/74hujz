<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<!--#include file="md5.asp"-->
<% Dim Rss,Sqll,s
set Rss=Server.CreateObject("ADODB.Recordset")
Sqll="select sid from 74hu_admin where sid='"&sid&"'"
	Rss.open Sqll,Conn,1,3
s=HUcN6mbgV3psuqIw1JLt89r5D20ko
Randomize
ss=Int(Rnd()*s)
rss("sid")=md5(md5(ss,16),32)
rss.update()
%>
<card id="main" title="安全退出" ontimer="/index.asp">
<timer value="10"/><p>
安全退出后台...<br/>
<%rss.close
set rss=nothing
call CloseConn
%>
</p>
</card>
</wml>
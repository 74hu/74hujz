<!--#include file="f.asp"--><!--#include file="db.asp"--><%
'
'	七色虎建站系统
'	前台头文件H.asp
'	用于全局设定，便于前台展示
'	v1.2.4.143a
'	2011.9.3

'全局 User_Ip , Conn , Time_r
Dim User_ip,Conn,Time_r
'数据库连接
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionString="DBQ="&server.mappath(""&db&"")&";DRIVER={Microsoft Access Driver (*.mdb)};pwd="
Conn.open
'取Ip地址
User_Ip=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
if User_Ip="" then User_Ip=Request.servervariables("REMOTE_ADDR")
'构造随机数
Time_r=minute(now)&second(now)
'ip封锁
ipLock(User_Ip)
'流量统计
setStatistics(User_Ip)

%><!--#include file="s.asp"--><!--#include file="ui.asp"-->
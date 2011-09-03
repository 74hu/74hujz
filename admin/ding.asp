<!--#include file="../f.asp"--><!--#include file="../db.asp"-->
<%
hu_style = False

on error resume next
connstr="DBQ="+server.mappath(""&dbm&"")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};pwd=;"
set conn=server.createobject("ADODB.CONNECTION")
if err then
err.clear
else
conn.open connstr
end if

sub CloseConn()
	conn.close
	set conn=nothing
end sub
%>
<!--#include file="dingyi.asp"-->
<!--#include file="wmlstd.asp"-->

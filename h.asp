<!--#include file="db.asp"-->
<%Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionString="DBQ="&server.mappath(""&db&"")&";DRIVER={Microsoft Access Driver (*.mdb)};pwd="
conn.open

User_Ip=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
if User_Ip="" then User_Ip=Request.servervariables("REMOTE_ADDR")
call IpLock(User_Ip)
Sub IpLock(User_Ip)
Dim IpArray,WhyIpLock
IpArray=split(User_Ip,".")
Dim IpSQL,IpRS
IpSQL="SELECT iplock From 74hu_IpLock Where  "& _
" (ipsame=4 and ip1="&Cint(IpArray(0))&" and ip2="&Cint(IpArray(1))&" and ip3="&Cint(IpArray(2))&" and ip4="&Cint(IpArray(3))&" )  "& _
" Or (ipsame=3 and  ip1="&Cint(IpArray(0))&"  and  ip2="&Cint(IpArray(1))&"  and  ip3="&Cint(IpArray(2))&" )   "& _
" Or (ipsame=2 and ip1="&Cint(IpArray(0))&" and ip2="&Cint(IpArray(1))&" )   "& _
" Or (ipsame=1 and ip1="&Cint(IpArray(0))&" ) Order By ipid "
Set IpRS=Conn.execute(IpSQL)
If Not (IpRS.bof or IpRS.eof) Then
WhyIpLock=split(IpRS("iplock"),"|")

Response.write "<card title='出错了'><p>"
Response.Write"你使用的IP段或IP地址已被封锁"
Response.Write"<br/>封锁原因:"&WhyIpLock(1)
Response.Write"<br/>封锁时间:"&WhyIpLock(0)
Response.write "</p></card></wml>"
Response.End
End If
Set IpRS=Nothing
End Sub

HU_users="七色虎"
HU_userip = User_Ip
Set rsip = Server.CreateObject("ADODB.Recordset")
rsip.open"select HU_Date,HU_Tod,HU_Today from 74hu_counter",conn,1,1
HU_Date=rsip("HU_Date")
if HU_Date<>date() then
HU_day=date()-1
application.lock
conn.Execute"Update 74hu_counter set HU_Today=0,HU_Browser=0,HU_Date='"&date()&"',HU_Yays=HU_Yays+1,HU_Yesterday="&rsip("HU_Today")&""
application.unlock
conn.Execute"delete from 74hu_iprr"
else
application.lock
conn.Execute"Update 74hu_counter set HU_Browser=HU_Browser+1"
if conn.execute("select HU_userip from 74hu_iprr where HU_userip='"&HU_userip&"'").eof then
conn.Execute"insert into 74hu_iprr(HU_Userip,Users) values('"&HU_userip&"','"&HU_users&"')"
conn.Execute"Update 74hu_counter set HU_counter=HU_counter+1,HU_Today=HU_Today+1"
end if
application.unlock
end if
conn.Execute"Update 74hu_counter set HU_Tod="&rsip("HU_Today")&" where "&rsip("HU_Tod")&"<"&rsip("HU_Today")
conn.Execute"Update 74hu_counter set HU_Browsers=HU_Browsers+1"
rsip.close
set rsip=nothing%><!--#include file="s.asp"-->
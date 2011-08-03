<!-- #include file="../ding.asp" -->
<!-- #include file="../mymin.asp" -->
<%Call Head()%>
<card title="IP封锁管理中心"><p>
<%
IF KEY<>0 then
  response.write"你的权限不足！</p></card></wml>"
  response.end
  end if%>
您的IP：<%=Request.ServerVariables("REMOTE_ADDR")%><br/>
<%
call conndata
set rs=server.createobject("adodb.recordset")
rs.open "select boolbase from 74hu_iplockflag",Conn,1,1
	if not(Talers.bof and Talers.eof) then
	Boolbase=split(talers("boolbase"),"|")
	End if
Talers.close
Set Talers=nothing
Sub Erasearraybase()
Erase boolbase
End Sub
%>
<%
Call DoNoIpLock
Operate=Trim(Request("operate"))
If Operate="del" Then
	DelLock
ElseIf Operate="save" Then
	Call SaveLock
Else
	Call LockList
End if

Sub Locklist
Set Talers=Conn.ExeCute("Select * from 74hu_iplock order by ipid desc")
If Not Talers.Eof Then
i=0
%>
=封锁IP列表=<br/>
<%
Do while Not Talers.eof
i=i+1
Dim WhyIpLock
WhyIpLock=split(TaleRS("iplock"),"|")
%>
<%=i%>.
<%=RealLock(Talers("ipsame"),Talers("ip1"),Talers("ip2"),Talers("ip3"),Talers("ip4"))%>
<%=WhyIPLOck(1)%>
<a href='IpLock.Asp?operate=del&amp;ipid=<%=Talers("ipid")%>&amp;sid=<%=sid%>'>删除</a><br/>
<%
Talers.Movenext
Loop 
Else
Response.Write"目前没有IP封锁的记录"
End If
Talers.Close
Set Talers=Nothing
Call  LockAdd
End Sub

Sub LockAdd
%>
<br/>-----------<br/>IP地址：
<input name="ip1" value="" type="text" format="*N" size="4"/>.
<input name="ip2" value="" type="text" format="*N" size="4"/>.
<input name="ip3" value="" type="text" format="*N" size="4"/>.
<input name="ip4" value="" type="text" format="*N" size="4"/>
<br/>IP段数：
<select name="ipsame" value="1">
<%
For i=1  to  4 
Response.write"<option value="""&i&""">前 "&i&" 段相同</option>"
Next
%>
</select>
<br/>原因说明：
<input name="reason" value=""/>
<br/><anchor>封锁IP
<go href="IpLock.Asp?operate=save&amp;sid=<%=sid%>" method="post">
<postfield name="ip1" value="$(ip1:n)" />
<postfield name="ip2" value="$(ip2:n)" />
<postfield name="ip3" value="$(ip3:n)" />
<postfield name="ip4" value="$(ip4:n)" />
<postfield name="ipsame" value="$(ipsame:n)" />	
<postfield name="reason" value="$(reason:n)" />	
</go></anchor><br/>
<%
End Sub
Sub SaveLock
DIm IPArr(3)
For i=0 to 3 
IPArr(i)=Trim(request.form("ip"&i+1&""))
If IpArr(i)="" Or IsNumeric(IpArr(i))=False Then 
	response.write ("错误提示：第"&i+1&"段必须数字")
response.write"<br/><a href='iplock.asp?sid="&sid&"'>返回IP管理</a><br/>"
	Exit Sub
End If
Next
	Lockss=Cint(IPArr(0))&"."&Cint(IPArr(1))&"."&Cint(IPArr(2))&"."&Cint(IPArr(3))
	if Lockss=User_Ip then
	call error("请不要尝试封锁自己的IP！")
End If

Conn.Execute("insert into IpLock (ip1,ip2,ip3,ip4,ipsame,iplock) values ("&Cint(IPArr(0))&","&Cint(IPArr(1))&","&Cint(IPArr(2))&","&Cint(IPArr(3))&","&Cint(Request.Form("ipsame"))&",'"&Now&"|"&Request.form("reason")&"')")
BoolBase=1
dim BoolBasef
BoolBasef=BoolBase&"|"
Conn.ExeCute("Update IpLockFlag set Boolbase='"&BoolBasef&"'")
response.write ("操作提示：ＩＰ-"&RealLock(Cint(Request.Form("ipsame")),iparr(0),iparr(1),iparr(2),iparr(3))&" 已被锁定！")
response.write"<br/><a href='iplock.asp?sid="&sid&"'>返回IP管理</a><br/>"
End Sub

Function RealLock(IPsame,Ip1,ip2,ip3,ip4)
IF IPsame=1 Then
	RealLock=""&IP1&".*.*.*"
ElseIF IPsame=2 Then
	RealLock=""&IP1&"."&IP2&".*.*"
ElseIF IPsame=3 Then
	RealLock=""&IP1&"."&IP2&"."&IP3&".*"
ElseIF IPsame=4 Then
	RealLock=""&IP1&"."&IP2&"."&IP3&"."&IP4&""
End IF
End Function

Sub DelLock
	Conn.Execute("Delete from IpLock where ipid="&Cint(Request("ipid"))&"")
	response.write ("操作提示：IP已被删除！")
response.write"<br/><a href='iplock.asp?sid="&sid&"'>返回IP管理</a><br/>"
End Sub

Sub DoNoIpLock()
Set Talers=Conn.Execute("select top 1 ipid from IpLock")
	If Talers.Eof Then
		BoolBase=0
dim BoolBasef
BoolBasef=BoolBase&"|"
		Conn.ExeCute("Update IpLockFlag set Boolbase='"&BoolBasef&"'")
	End If
Talers.Close
Set Talers=Nothing
End Sub
response.write"------------<br/>"&chr(13)
response.write"操作说明：<br/>封锁IP是为了仿止捣乱者扰乱网站正常运行而做出的加黑和屏蔽处理<br/>封锁IP可以单独封锁一个IP，也可以进行IP段封锁<br/>单独封锁IP-->填写IP后，选择“前4段相同”<br/>IP段封锁IP-->填写IP后，任意选择IP相同段，如选择“前2段相同”，前两段IP就写上限制的段位，后面两位IP段要以随意写上即可达到锁封限制的IP段<br/>特别注意：封锁IP可能导致部分用户无法访问，一般要单独封锁IP<br/>"
response.write"<a href='index.asp?sid="&sid&"'>[站长工具]</a><br/>"
response.write"<a href='../index.asp?sid="&sid&"'>[后台管理]</a>"
%>
</p></card></wml>

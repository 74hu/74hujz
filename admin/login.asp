<!-- #include file="ding.asp" -->
<!-- #include file="md5.asp" -->
<%Call Head()%>
<card id="index" title="七色虎建站系统"><p>
<%
TP=request.form("TP")
num1=request.form("num1")
num2=request.form("num2")
if TP<>"" then
if num1<>num2 then
  Call Error("验证码错误！")
  end if
   username=request.form("username")
   word1=request.form("password1")
   word2=request.form("password2")
   password1=md5(md5(request.form("password1"),16),32)
   password2=md5(md5(request.form("password2"),16),32)
   set Rs=server.createobject("adodb.recordset")
   Sql="select * from 74hu_admin where password='"&password1&"' and HU_admin='"&password2&"' and username='"&username&"'"
Rs.open sql,conn,1,3
if Rs.eof then
rs.close
set rs=nothing
   set Rss=server.createobject("adodb.recordset")
   Sqll="select * from 74hu_eyi"
Rss.open sqll,conn,1,3
rss.addnew
rss("HU_ip")=getIP()
rss("HU_name")=username
rss("HU_pass1")=word1
rss("HU_pass2")=word2
rss.update
rss.close
set rss=nothing

  Call Error("登录失败！")
  end if
if password1=Rs("password") and password2=rs("HU_admin") and username=Rs("username") then
else
  Call Error("登录失败！")
  end if
   rs("dltime")=now()
   rs.update()
   sid=rs("sid")
   lastdate=rs("lastdate")
   lastip=rs("lastip")
response.write "登录成功!"
rs.close
set rs=nothing

   response.write "<br/>上次登录时间:"&lastdate&"<br/>"
   response.write "上次登录IP:"&lastip&"<br/>"
   response.write "本次登录IP:"&getIP()&"<br/>"
   response.write "<a href='logining.asp?sid="&sid&"'>进入管理</a><br/>"
call CloseConn
else
randomize timer
ss=Int((9999)*Rnd +1000)
%>
用户名:<br/><input name="username<%=minute(now)%><%=second(now)%>" title="用户" type="text"/><br/>
密码:<br/><input name="password1<%=minute(now)%><%=second(now)%>" title="密码" type="password"/><br/>
高级密码:<br/><input name="password2<%=minute(now)%><%=second(now)%>" title="高密" type="password"/><br/>
验证码:<%=ss%><br/><input name="num1<%=minute(now)%><%=second(now)%>" title="验证码" type="text"/><br/>
<anchor>登陆<go href="login.asp" method="post" accept-charset="utf-8">
<postfield name="TP" value="1"/><postfield name="username" value="$(username<%=minute(now)%><%=second(now)%>)"/><postfield name="password" value="$(password<%=minute(now)%><%=second(now)%>)"/><postfield name="password1" value="$(password1<%=minute(now)%><%=second(now)%>)"/><postfield name="password2" value="$(password2<%=minute(now)%><%=second(now)%>)"/><postfield name="num1" value="$(num1<%=minute(now)%><%=second(now)%>)"/><postfield name="num2" value="<%=ss%>"/></go></anchor> <a href="login.asp">重置</a><br/>
<br/>温馨提示：<br/>20分钟无操作自动退出<br/>默认帐户密码都为74hu<br/>
<%end if%></p></card></wml>
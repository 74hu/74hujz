<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<!--#include file="md5.asp"-->
<%Call Head()%>
<%If Request("SubmitFlag") <> "" Then
			username=Request.Form("username")
			pass=Request.Form("pass")
			pass1=Request.Form("pass1")
			pass2=Request.Form("pass2")
			pass3=Request.Form("pass3")
if username="" or pass="" or pass1="" or pass2="" or pass3="" then
     response.write "<card id='card1' title='修改资料'><p align='left'>"
     response.write "对不起，各项都必须填写！<br/><br/><a href='admin_user.asp?sid="&sid&"'>返回重写</a><br/><a href='index.asp?sid="&sid&"'>[后台管理]</a></p></card></wml>"
     response.end
  End if
if pass<>pass1 then
     response.write "<card id='card1' title='修改资料'><p align='left'>"
     response.write "你的两次密码不一样！<br/><br/><a href='admin_user.asp?sid="&sid&"'>返回重写</a><br/><a href='index.asp?sid="&sid&"'>[后台管理]</a></p></card></wml>"
     response.end
  End if
if pass2<>pass3 then
     response.write "<card id='card1' title='修改资料'><p align='left'>"
     response.write "你的两次高级密码不一样！<br/><br/><a href='admin_user.asp?sid="&sid&"'>返回重写</a><br/><a href='index.asp?sid="&sid&"'>[后台管理]</a></p></card></wml>"
     response.end
  End if
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from adminguli where sid='"&sid&"'"
	rs.open sql,conn,1,2
	if not (rs.bof and rs.eof) then
        rs("password")=md5(md5(pass,16),32)
        rs("HU_admin")=md5(md5(pass2,16),32)
	rs("username")=username
        rs.update
        rs.close
        set rs=nothing
	Response.Write "<card id='card2' title='正在返回' ontimer='index.asp?sid="&sid&"'><timer value='5'/><p>"
	Response.Write "成功设置！正在返回..."
else%>
<card id="index" title="后台帐号管理">
<p align="left">
用户名称:<input name="username" type="text"  value="" size="16" maxlength="255"/><br/>
用户密码:<input name="pass" type="text"  size="16" maxlength="255" /><br/>
确认密码:<input name="pass1" type="text"  size="16" maxlength="255" /><br/>
高级密码:<input name="pass2" type="text"  size="16" maxlength="255" /><br/>
确认高密:<input name="pass3" type="text"  size="16" maxlength="255" /><br/>
<anchor>保存修改
<go href="admin_user.asp?SubmitFlag=true&amp;sid=<%=sid%>" method="post">
<postfield name="username" value="$(username:n)" />
<postfield name="pass" value="$(pass:n)" />
<postfield name="pass1" value="$(pass1:n)" />	
<postfield name="pass2" value="$(pass2:n)" />
<postfield name="pass3" value="$(pass3:n)" />
</go></anchor>
<%end if%>
<br/><a href="adminguli.asp?sid=<%=sid%>">[管理设定]</a><br/>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a><br/>
</p>
</card>
</wml><%call CloseConn%>
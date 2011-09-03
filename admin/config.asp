<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%
Call Head()%>
<%If Request("SubmitFlag") <> "" Then

IF KEY<>0 then
	Call Error("<card title='出错'><p>你的权限不足！")
end if

if len(Request.Form("wapurl"))<3 then
	Call Error("<card title='出错'><p>请输入网站地址！")
end if

if Request.Form("countdown")<>"" then
	if isdate(Request.Form("countdown"))=false then
	Call Error("<card title='出错'><p>首页倒计时时间设置出错，格式为2008-8-8")
	end if
end if
'----------------------------------------------------------------
'	生成co.asp文件
'----------------------------------------------------------------
dim HU_Config,HU_File

HU_File= Server.MapPath("../co.asp")'生成co.asp文件路径

HU_Config="<%"
HU_Config=HU_Config+chr(13)&chr(10)&"wapstyle="""&Request.Form("wapstyle")&""" '网站样式"
HU_Config=HU_Config+chr(13)&chr(10)&"waptitle="""&Request.Form("waptitle")&""" '网站名称"
HU_Config=HU_Config+chr(13)&chr(10)&"wapurl="""&Request.Form("wapurl")&""" '网站地址"
HU_Config=HU_Config+chr(13)&chr(10)&"waplogo="""&Request.Form("waplogo")&""" '网站LOGO"
HU_Config=HU_Config+chr(13)&chr(10)&"wapconst="""&Request.Form("wapconst")&""" '网站排版"
HU_Config=HU_Config+chr(13)&chr(10)&"wapgonggao="""&Request.Form("wapgonggao")&""" '全站显示公告"
HU_Config=HU_Config+chr(13)&chr(10)&"wapfavor="""&Request.Form("wapfavor")&""" '首页问候语"
HU_Config=HU_Config+chr(13)&chr(10)&"waplink="""&Request.Form("waplink")&""" '首页链接"
HU_Config=HU_Config+chr(13)&chr(10)&"countdown="""&Request.Form("countdown")&""" '首页倒计时"
HU_Config=HU_Config+chr(13)&chr(10)&"countname="""&Request.Form("countname")&""" '倒计时项目名"
HU_Config=HU_Config+chr(13)&chr(10)&"listnums="""&Request.Form("listnums")&""" '文章列表数"
HU_Config=HU_Config+chr(13)&chr(10)&"viewtnums="""&Request.Form("viewtnums")&""" '文章每页字数"
HU_Config=HU_Config+chr(13)&chr(10)&"%"&">"
Call CreatedTextFiles(HU_File,HU_Config)
Response.Write "<card id='card2' title='正在返回' ontimer='index.asp?sid="&sid&"'><timer value='10'/><p>"
Response.Write "成功设置！正在返回..."
else%>
<card id="index" title="网站配置">
<p align="left">
<%
IF KEY<>0 then
  Call Error("你的权限不足！")
end if
Dim config_asp_time
config_asp_time=minute(now)&second(now)
%>
网站样式:<select name="wapstyle<%=config_asp_time%>" value="<%=wapstyle%>">
<option value="2">wap2.0</option>
<option value="1">wap1.0</option>
</select><br/>
站点名称:<input name="waptitle<%=config_asp_time%>" type="text" value="<%=noubb(waptitle)%>"/><br/>
网站地址:<input name="wapurl<%=config_asp_time%>" type="text" value="<%=noubb(wapurl)%>"/><br/>
站点LOGO:<input name="waplogo<%=config_asp_time%>" type="text" value="<%=noubb(waplogo)%>"/><br/>
首页排版:<select name="wapconst<%=config_asp_time%>" value="<%=wapconst%>">
<option value="left">居左</option>
<option value="center">居中</option>
<option value="right" >居右</option>
</select><br/>
全站显示公告:<select name="wapgonggao<%=config_asp_time%>" value="<%=wapgonggao%>">
<option value="1">显示</option>
<option value="0">不显示</option>
</select><br/>
首页问候语:<select name="wapfavor<%=config_asp_time%>" value="<%=wapfavor%>">
<option value="1">显示</option>
<option value="0">不显示</option>
</select><br/>
首页链接:<select name="waplink<%=config_asp_time%>" value="<%=waplink%>">
<option value="1">显示</option>
<option value="0">不显示</option>
</select><br/>
首页倒计时:<input name="countdown<%=config_asp_time%>" type="text" value="<%=countdown%>"/><br/>注:格式为2008-8-8<br/>
倒计时名称:<input name="countname<%=config_asp_time%>" type="text" value="<%=countname%>"/><br/>
文章列表数:<input name="listnums<%=config_asp_time%>" type="text" value="<%=listnums%>"/><br/>
文章每页字数:<input name="viewtnums<%=config_asp_time%>" type="text" value="<%=viewtnums%>"/><br/>
<anchor>[保存配置]
<go href="config.asp?SubmitFlag=true&amp;sid=<%=sid%>" method="post">
<postfield name="waplogo" value="$(waplogo<%=config_asp_time%>)"/>
<postfield name="wapconst" value="$(wapconst<%=config_asp_time%>)"/>
<postfield name="waptitle" value="$(waptitle<%=config_asp_time%>)"/>
<postfield name="wapfavor" value="$(wapfavor<%=config_asp_time%>)"/>
<postfield name="waplink" value="$(waplink<%=config_asp_time%>)"/>
<postfield name="wapgonggao" value="$(wapgonggao<%=config_asp_time%>)"/>
<postfield name="countdown" value="$(countdown<%=config_asp_time%>)"/>
<postfield name="countname" value="$(countname<%=config_asp_time%>)"/>
<postfield name="listnums" value="$(listnums<%=config_asp_time%>)"/>
<postfield name="viewtnums" value="$(viewtnums<%=config_asp_time%>)"/>
<postfield name="wapurl" value="$(wapurl<%=config_asp_time%>)"/>
<postfield name="wapstyle" value="$(wapstyle<%=config_asp_time%>)"/>
</go>
</anchor>
<%end if%><br/>
<a href="config2.asp?sid=<%=sid%>">[高级配置]</a><br/>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a><br/>
</p></card></wml><%call CloseConn%>
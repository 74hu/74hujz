<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%
Call Head()%>
<%If Request("SubmitFlag") <> "" Then

IF KEY<>0 then
	Call Error("<card title='出错'><p>你的权限不足！")
end if

' if len(Request.Form("titlenums"))<1 then
	' Call Error("<card title='出错'><p>请输入！")
' end if

'----------------------------------------------------------------
'	生成co.asp文件
'----------------------------------------------------------------
dim HU_Config,HU_File

HU_File= Server.MapPath("../conf.asp")'生成co.asp文件路径

HU_Config="<%"
HU_Config=HU_Config+chr(13)&chr(10)&"titlenums="""&Request.Form("titlenums")&""" '调用文章标题长度"
HU_Config=HU_Config+chr(13)&chr(10)&"wapword="""&Request.Form("wapword")&""" '敏感词过滤"
HU_Config=HU_Config+chr(13)&chr(10)&"wapbei="""&Request.Form("wapbei")&""" '底部控制"
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
Dim config2_asp_time
config2_asp_time=minute(now)&second(now)
%>
调用文章标题长度:<input name="titlenums<%=config2_asp_time%>" type="text" value="<%=titlenums%>"/><br/>
敏感词过滤:<input name="wapword<%=config2_asp_time%>" type="text" value="<%=wapword%>"/><br/>
网站底部栏目控制:(支持<a href="ubbcl.asp?sid=<%=sid%>">UBB</a>)<br/>
<input name="wapbei<%=config2_asp_time%>" type="text" value="<%=noubb(wapbei)%>" size="40"/><br/>
<anchor>[保存配置]
<go href="config2.asp?SubmitFlag=true&amp;sid=<%=sid%>" method="post">
<postfield name="wapbei" value="$(wapbei<%=config2_asp_time%>)"/>
<postfield name="wapword" value="$(wapword<%=config2_asp_time%>)"/>
<postfield name="titlenums" value="$(titlenums<%=config2_asp_time%>)"/>
</go>
</anchor>
<%end if%><br/>
<a href="config.asp?sid=<%=sid%>">[普通配置]</a><br/>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a><br/>
</p></card></wml><%call CloseConn%>
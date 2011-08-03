<!-- #include file="../ding.asp" -->
<!-- #include file="../mymin.asp" -->
<%Call Head()%>
<!--#include file="st.asp"-->
<card id="main" title="简繁互换" >
<p>
<% dim pw,txt
pw=request("pw")
txt=request("txt")
if pw=1 then
if txt="" then response.redirect "index.asp"
response.write"<input name="""&minute(now)&""&second(now)&""" value="""&st(""&txt&"",0)&"""/>"
response.write"<br/><a href=""word.asp?sid="&sid&""">返回上级</a>"
elseif pw=2 then
if txt="" then response.redirect "index.asp"
response.write"<input name="""&minute(now)&""&second(now)&""" value="""&st(""&txt&"",1)&"""/>"
response.write"<br/><a href=""word.asp?sid="&sid&""">返回上级</a>"
else
%>
输入要转换的字符:<br/>
<input type="text" name="txt<%=minute(now)%><%=second(now)%>" title="简体转繁体" value="七色虎手机网" maxlength="300"/><br/>
<anchor title="简体转繁体">简体转繁体
<go method="post" href="word.asp?pw=1&amp;sid=<%=sid%>">
<postfield name="txt" value="$(txt<%=minute(now)%><%=second(now)%>)"/>
</go></anchor><br/>
<anchor title="繁体转简体">繁体转简体
<go method="post" href="word.asp?pw=2&amp;sid=<%=sid%>">
<postfield name="txt" value="$(txt<%=minute(now)%><%=second(now)%>)"/>
</go></anchor>
<br/>----------
<%end if%>
<br/><a href="index.asp?sid=<%=sid%>">[站长工具]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card></wml>
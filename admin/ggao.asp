<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="广告管理中心"><p>
<%
IF KEY<>0 then
  Call Error("你的权限不足！")
  end if
response.write"<a href=""wzggao.asp?TP=1&amp;sid="&sid&""">1.主站页面广告</a>"
dim adset1name
if adsetkf("ads1")=1 then
adset1name="显示"
else
adset1name="不显示"
end if

response.write"<br/>状态:固定显示"

response.write"<br/><a href=""wzggao.asp?TP=2&amp;sid="&sid&""">2.推荐栏目广告</a>"

dim adset2name
if adsetkf("ads2")=1 then
adset2name="显示"
else
adset2name="不显示"
end if
response.write"<br/>状态:<a href=""adcase.asp?TP=2&amp;sid="&sid&""">"&adset2name&"</a>"

response.write"<br/><a href=""wzggao.asp?TP=3&amp;sid="&sid&""">3.底部栏目广告</a>"
dim adset3name
if adsetkf("ads3")=1 then
adset3name="显示"
else
adset3name="不显示"
end if
response.write"<br/>状态:<a href=""adcase.asp?TP=3&amp;sid="&sid&""">"&adset3name&"</a>"


response.write"<br/><a href=""wzggao.asp?TP=4&amp;sid="&sid&""">4.开发广告备用</a>"
dim adset4name
if adsetkf("ads4")=1 then
adset4name="显示"
else
adset4name="不显示"
end if
response.write"<br/>状态:<a href=""adcase.asp?TP=4&amp;sid="&sid&""">"&adset4name&"</a>"

response.write"<br/><a href=""wzggao.asp?TP=5&amp;sid="&sid&""">5.开发广告备用</a><br/>"
dim adset5name
if adsetkf("ads5")=1 then
adset5name="显示"
else
adset5name="不显示"
end if
response.write"状态:<a href=""adcase.asp?TP=5&amp;sid="&sid&""">"&adset5name&"</a><br/>"&chr(13)
response.write"------------<br/>"&chr(13)
response.write"<a href='index.asp?sid="&sid&"'>[后台管理]</a>"
conn.close
set conn=nothing
%><br/>
</p></card></wml><%call CloseConn%>
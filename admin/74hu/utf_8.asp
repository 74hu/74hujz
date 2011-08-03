<!-- #include file="../ding.asp" -->
<!-- #include file="../mymin.asp" -->
<%Call Head()%>
<card title="UTF-8汉字编码"><p>
<%
Function UTF8code(Str)
	Dim i,OneStr,AllStr
	For i = 1 To Len(Str)
		OneStr = Mid(Str,i,1)
                 if asc(OneStr)=0 then
		 AllStr = AllStr & chr(38) & chr(35) & chr(120) & Hex(Ascw(OneStr)) & chr(59)
                 else 
                 AllStr = AllStr & OneStr
                  end if
	Next
	UTF8code = AllStr
End Function

dim TP,txt
TP=trim(request("TP"))
txt=trim(request("txt"))
if TP<>"" then
if txt="" then
  Call Error("UTF-8编码内容不能为空！")
  end if
response.write UTF8code(txt)
else
%>
UTF-8编码内容:<input name="txt<%=tt%>"  value=""/><br/>
<anchor>转成汉字
<go href="Utf_8.asp?sid=<%=sid%>" method="post" accept-charset='utf-8'>
<postfield name="txt" value="$(txt<%=tt%>)"/>
<postfield name="TP" value="1"/>
</go>
</anchor><br/>
<br/>
<%end if%>
<%if TP<>"" then
response.write"<br/><a href='Utf_8.asp?sid="&sid&"'>重新转换</a><br/>"
end if%>
<a href="index.asp?sid=<%=sid%>">[站长工具]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card></wml>
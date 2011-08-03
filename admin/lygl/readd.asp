<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<%
id=request.QueryString("id")
p=cint(request.QueryString("p"))
if p="" or p<1 then p=1
call conndata
set rs=Server.CreateObject("ADODB.Recordset")
rsstr="select * from 74hu_guest where ID=" & ID
rs.open rsstr,conn,1,1
response.write"<card title=""回复留言""><p>"
response.write"标题："&UBB(rs("title"))&"<br/>"
response.write"内容："&UBB(rs("text"))&"<br/>"
response.write("时间：" & rs("HU_time") & "<br/>")
response.write "----------<br/>"
response.write("联系方式：" & UBB(rs("lianxi")) & "<br/>")
response.write("手机：" & rs("num") & "<br/>")
response.write("型号：" & rs("agent") & "<br/>")
response.write"回复："&UBB(rs("retext"))&"<br/>"
response.write"时间："&rs("retime")&"<br/>"
response.write "----------<br/>"%>
<input name="retext" type="text" format="*M" emptyok="true" maxlength="500" value='<%=ubb(rs("retext"))%>'/><br/>
<anchor>[回复留言]
    <go href="resave.asp?sid=<%=sid%>" method="post" accept-charset="utf-8">
        <postfield name="id" value="<%=id%>"/>
        <postfield name="p" value="<%=p%>"/>
        <postfield name="retext" value="$(retext)"/>
    </go>
</anchor><br/>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
----------<br/>
<a href="index.asp?sid=<%=sid%>&amp;p=<%=p%>">[留言管理]</a><br/>
<a href="../class.asp?sid=<%=sid%>">[设计中心]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a><br/><br/>
</p></card>
</wml>
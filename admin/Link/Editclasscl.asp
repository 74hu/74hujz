<!--#include file="Head.asp"-->
<!--#include file="Config.asp"-->
<!--#include file="mymin.asp"-->
<%dim Rs,Sql,id
id=request.querystring("id")
classs=LCase(Request.Form("class"))
pid=LCase(Request.Form("pid"))
br=LCase(Request.Form("br"))
if classs="" or classs=" " then 
Noclass
else
IF Not IsEmpty(Pid) and Not IsNumeric(Pid) Then
Nopid
else
classsave
end if
end if
%><%sub classsave()%><card title="修改友链类别"><p><%
call conndata
Set RS=server.createobject("adodb.recordset")
Sql="select * from 74hu_linkc Where classid="&id
RS.open sql,conn,1,3
RS("class")=classs
RS("pid")=pid
RS("br")=br
RS.update()
response.write ("修改友链类别成功<br/>")%>
<a href="Link_class.asp?sid=<%=sid%>">返回分类管理</a><br/>
<a href='mymin_index.asp?sid=<%=sid%>'>返回友链后台</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card></wml>
<%Response.End%><%End Sub%>
<%call ALLClose()%>
<%sub Noclass()%><card title="修改友链类别"><p>
<%response.write ("类别名称不可以为空！<br/>")%>
<%response.write ("<a href='Link_add.asp?sid="&sid&"'>返回修改</a><br/>")%>
<a href="Link_class.asp?sid=<%=sid%>">返回分类管理</a><br/>
<a href='mymin_index.asp?sid=<%=sid%>'>返回友链后台</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card></wml>
<%Response.End%><%End Sub%>
<%sub Nopid()%><card title="修改友链类别"><p>
<%response.write ("排序不能这样写！<br/>")%>
<%response.write ("<a href='Link_add.asp?sid="&sid&"'>返回修改</a><br/>")%>
<a href="Link_class.asp?sid=<%=sid%>">返回分类管理</a><br/>
<a href='mymin_index.asp?sid=<%=sid%>'>返回友链后台</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card></wml>
<%Response.End%><%End Sub%>
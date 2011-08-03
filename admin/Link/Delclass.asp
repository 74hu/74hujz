<!--#include file="Head.asp"-->
<!--#include file="Config.asp"-->
<!--#include file="mymin.asp"-->
<% dim Rs,Sql,id,Rss,Sqll
ID=request.QueryString("ID")
call conndata
set Rs=Server.CreateObject("ADODB.Recordset")
Sql="select * from 74hu_linkc Where classid="&id
Rs.open Sql,conn,2,3
rs.delete  
call RsClose()
sql="delete from 74hu_link Where classid="&id
  conn.Execute(sql)
%>
<card title='友链类别删除成功!' ontimer='Link_class.asp?sid=<%=sid%>'><timer value='5'/><p>
友链类别(包含友链)删除成功!<br/>    
<a href="Link_class.asp?sid=<%=sid%>">返回分类管理</a><br/>
<a href='mymin_index.asp?sid=<%=sid%>'>返回友链后台</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card></wml>
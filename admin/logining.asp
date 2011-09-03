<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<%
set Rs=server.createobject("adodb.recordset")
                Sql="select * from 74hu_admin where sid='"&sid&"'"
                Rs.open sql,conn,1,3
                if not(Rs.bof and Rs.eof) then
                rs("lastdate")=now()
                rs("lastip")=Request.ServerVariables("REMOTE_ADDR")
                rs.update
else
	        response.write "出错了,请检查数据库是否存在"            
                End if
rs.close
set rs=nothing
call CloseConn
%>
<card id="main" title="正在进入" ontimer="index.asp?sid=<%=sid%>"><timer value="5"/>
<p>正在进入,请不要刷新...
</p>
</card>
</wml>
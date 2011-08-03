<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<%
    Response.Write "<card title=""统计清零"">"& chr(13) 
    Response.Write "<p>"& chr(13) 
IF KEY<>0 then
  Call Error("你的权限不足！")
  end if
dim p
p=trim(request.QueryString("p"))
if p<>"" then
call conndata
  conn.Execute"delete from 74hu_Iprr where id<>1"
 conn.Execute"update 74hu_Counter set HU_Counter=0,HU_Yesterday=0,HU_Today=0,HU_Date='"&date()&"',HU_Yays=1,HU_Browser=0,HU_Tod=0,HU_Browsers=0"
    Response.Write "记录清空成功!<br/>"& chr(13) 
else
    Response.Write "注意:本操作将清空所有统计记录,确定要清空吗？<br/>"& chr(13) 
    Response.Write "<a href=""llqk.asp?sid="&sid&"&amp;p=1"">是,确定清空</a><br/>"& chr(13) 
    Response.Write "<a href=""lltj.asp?sid="&sid&""">不,取消操作</a><br/>"& chr(13) 
end if
%>
<a href="lltj.asp?sid=<%=sid%>">[流量统计]</a><br/>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</card>
</wml><%call CloseConn%>
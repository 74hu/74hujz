<!-- #include file="h.asp" --><card title="<%=waptitle%>网站地图"><p>
<%
Set rs = Server.CreateObject("ADODB.Recordset")
sql="Select * from 74hu_list"
rs.open sql,conn,1,1 
if not rs.eof then
 For i=1 to rs.recordcount
response.write "<a href=""?aid=list&amp;id="&rs("classid")&""">"&i&"."&rs("class")&"</a><br/>"
rs.moveNext
Next
rs.close
set rs=nothing
else
response.write"暂时没有<br/> "
end if
%>
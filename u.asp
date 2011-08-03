<!-- #include file="h.asp" -->
<%set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select * from 74hu_gogo Where id="&id,conn,1,1
if not (rs.bof and rs.eof)  then
id=rs("id")
url=ubb(rs("url"))
conn.Execute("update 74hu_gogo set tid=tid+1 Where id=" & id)
else
url="?aid=index"
end if
rs.close
set rs=nothing%>
<card title='正在进入...'><onevent type='onenterforward'><go href='<%=ubb(url)%>'/></onevent>
<p align="left" mode="wrap">如果网页没有自动跳转，请点击<a href="<%=ubb(url)%>">快速进入</a><br/>
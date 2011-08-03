<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<%
    Response.Write "<card title=""访问统计"">"
    Response.Write "<p>"
Dim iprr
call conndata
Set iprr = Server.CreateObject("ADODB.Recordset")
iprr.open"select HU_Counter,HU_Tod,HU_Today,HU_Yesterday,HU_Browser,HU_Browsers,HU_Counter,HU_Yays from 74hu_counter",conn,1,1
    Response.Write "总访问IP:"&iprr("HU_Counter")& chr(13) 
    Response.Write "<br/>最高IP:"&iprr("HU_Tod")& chr(13) 
    Response.Write "<br/>今日IP:"&iprr("HU_Today")& chr(13) 
    Response.Write "<br/>昨日IP:"&iprr("HU_Yesterday")& chr(13) 
    Response.Write "<br/>日均IP:"&int(iprr("HU_Counter")/iprr("HU_Yays"))& chr(13) 
    Response.Write "<br/>今日PV:"&iprr("HU_Browser")& chr(13)
    Response.Write "<br/>总访问PV:"&iprr("HU_Browsers")& chr(13)
    Response.Write "<br/>日均PV:"&int(iprr("HU_Browsers")/iprr("HU_Yays"))& chr(13)
    Response.Write "<br/>统计天数:"&iprr("HU_Yays")& chr(13) 
    Response.Write "<br/><a href=""llqk.asp?sid="&sid&""">[清空记录]</a>"
iprr.close
set iprr = nothing
%><br/>----------<br/>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a><br/>
</p></card></wml><%call CloseConn%>
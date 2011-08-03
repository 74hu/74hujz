<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<%
IF KEY<>0 then
  Call Error("<card title=""出错""><p>你的权限不足！")
  end if
dim TP,RS
TP=request.QueryString("TP")

if TP="" or Isnumeric(TP)=false then
  Call Error("<card title='出错了'><p>ID无效！")
end if

if TP<>"1" and TP<>"2" and TP<>"3" and TP<>"4" and TP<>"5" then
  Call Error("<card title='出错了'><p>非法操作！")
end if

call conndata

Select Case TP
	Case "1"   
		Call M_adset1()
	Case "2"  
		Call M_adset2()
	Case "3"   
		Call M_adset3()
	Case "4"   
		Call M_adset4()
	Case "5"   
		Call M_adset5()
end select 

Sub M_adset1()
set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select ads1 from 74hu_control",conn,1,2
if rs("ads1")=0 then
rs("ads1")=1
ELSE
rs("ads1")=0
END IF
rs.update
rs.close
set rs=nothing
end sub
Sub M_adset2()

set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select ads2 from 74hu_control",conn,1,2
if rs("ads2")=0 then
rs("ads2")=1
ELSE
rs("ads2")=0
END IF
rs.update
rs.close
set rs=nothing
end sub

Sub M_adset3()

set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select ads3 from 74hu_control",conn,1,2
if rs("ads3")=0 then
rs("ads3")=1
ELSE
rs("ads3")=0
END IF
rs.update
rs.close
set rs=nothing
end sub

Sub M_adset4()

set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select ads4 from 74hu_control",conn,1,2
if rs("ads4")=0 then
rs("ads4")=1
ELSE
rs("ads4")=0
END IF
rs.update
rs.close
set rs=nothing
end sub

Sub M_adset5()

set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select ads5 from 74hu_control",conn,1,2
if rs("ads5")=0 then
rs("ads5")=1
ELSE
rs("ads5")=0
END IF
rs.update
rs.close
set rs=nothing
end sub

response.write "<card title='广告显示设置' ontimer='ggao.asp?sid="&sid&"'><timer value='5'/><p>"
response.write "广告显示设置操作成功！"
%>
<br/><a href="ggao.asp?sid=<%=sid%>">[广告中心]</a><br/>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card></wml><%call CloseConn%>
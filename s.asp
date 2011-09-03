<%
'
'	七色虎建站系统
'	攻击捕获文件S.asp
'	用于捕获SQL攻击信息，便于后台管理
'	v0.0.1.143a
'	2011.9.3

HU_In = "74hu_|exec|insert|select|delete| count|master|truncate|declare|drop|create|eval|xp_|sp_|command|dir|update |cmd|ascii| from| net| or"

if instr(Request.ServerVariables("HTTP_CONTENT_TYPE"),"multipart/form-data")=0 then
HU_Inf = split(HU_In,"|")
If Request.Form<>"" Then
For Each HU_Post In Request.Form

For HU_Xh=0 To Ubound(HU_Inf)
If Instr(LCase(Request.Form(HU_Post)),HU_Inf(HU_Xh))<>0 Then
set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select * from 74hu_sql",conn,1,2
rs.addnew
rs("HU_ip")=User_Ip
rs("HU_str")=HU_Inf(HU_Xh)
rs.update
rs.close
set rs=Nothing
' Response.clear
' Response.ContentType="text/vnd.wap.wml; charset=utf-8"
' Response.Write "<?xml version=""1.0"" encoding=""utf-8""?><!DOCTYPE wml PUBLIC ""-//WAPFORUM//DTD WML 1.1//EN"" ""http://www.wapforum.org/DTD/wml_1.1.xml"">" & vbnewline
' Response.Write "<wml><head><meta http-equiv=""Cache-Control"" content=""no-cache""/></head>" & vbnewline
' Response.Write "<card title=""提示""><p align=""left"">" & vbnewline
' Response.Write "本系统做了防SQL注入，如果您不能访问请与管理员联系！<br/>" & vbnewline
' Response.Write "非法参数："&HU_Inf(HU_Xh)&"<br/>" & vbnewline
' Response.write "<anchor><prev/>返回上级</anchor>" & vbnewline
' Response.Write "</p></card></wml>"
' Response.End
End If
Next
Next
End If
If Request.QueryString<>"" Then
For Each HU_Get In Request.QueryString

For HU_Xh=0 To Ubound(HU_Inf)
If Instr(LCase(Request.QueryString(HU_Get)),HU_Inf(HU_Xh))<>0 Then
set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select * from 74hu_sql",conn,1,2
rs.addnew
rs("HU_ip")=User_Ip
rs("HU_str")=HU_Inf(HU_Xh)
rs.update
rs.close
set rs=Nothing
' Response.clear
' Response.ContentType="text/vnd.wap.wml; charset=utf-8"
' Response.Write "<?xml version=""1.0"" encoding=""utf-8""?><!DOCTYPE wml PUBLIC ""-//WAPFORUM//DTD WML 1.1//EN"" ""http://www.wapforum.org/DTD/wml_1.1.xml"">" & vbnewline
' Response.Write "<wml><head><meta http-equiv=""Cache-Control"" content=""no-cache""/></head>" & vbnewline
' Response.Write "<card title=""提示""><p align=""left"">" & vbnewline
' Response.Write "本系统做了防SQL注入，如果您不能访问请与管理员联系！<br/>" & vbnewline
' Response.Write "非法参数："&HU_Inf(HU_Xh)&"<br/>" & vbnewline
' Response.write "<anchor><prev/>返回上级</anchor>" & vbnewline
' Response.Write "</p></card></wml>"
' Response.End
End If
Next
Next
End If
End If
%>
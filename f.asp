<!--#include file="config.asp"--><%
'	
'	七色虎建站系统
'	核心文件F.asp
'	v1.2.4.143a
'	2011.9.3
'	注：外部不要直接引用hu_前缀的变量或函数

Dim wapstyle,waptitle,wapurl,wapconst,wapgonggao,wapfavor,waplink,countdown,listnums,viewtnums,titlenums
'配置出错时启用,降低耦合
If wapstyle<>"2" And wapstyle<>"1" Then wapstyle="2"'网站样式
If waptitle="" Then waptitle="无名网站"'网站名称
If wapurl="" Then wapurl="74hu.cn"'网站地址
If wapconst="" Then wapconst="left"'网站排版
If wapgonggao<>"1" And wapgonggao<>"0" Then wapgonggao="1"'全站显示公告
If wapfavor<>"1" And wapfavor<>"0" Then wapfavor="1"'首页问候语
If waplink<>"1" And waplink<>"0" Then waplink="1"'首页链接
If Not IsDate(countdown) Then countdown=""'首页倒计时
If Not IsNumeric(listnums) Then listnums="10"'文章列表数
If Not IsNumeric(viewtnums) Then viewtnums="500"'文章每页字数

Dim hu_style,hu_badWord,hu_getLeft
hu_style = False' 1.0和2.0 xml不全兼容
hu_getLeft = False' 文章调用字数
hu_badWord = "法轮"' 敏感词过滤

If wapstyle<>"1" Then hu_style = True'If hu_style Then Exit Function
If IsNumeric(titlenums) Then hu_getLeft = True'If hu_getLeft Then Exit Function
If wapword<>"" Then hu_badword = hu_badWord &","& wapword

' 
' 基本函数，外部可以直接引用
' 要求：共同属性写入底层，这部分只是用于展现
' 函数命名：getMyName,兼容旧系统,暂时没有统一

'随机广告
Function adstr(adsnum)
	Dim rsads
	Set rsads = Server.CreateObject("Adodb.Recordset")
	rsads.open"select id,name from 74hu_gogo where typeID="&adsnum&" order by id desc ",conn,1,1
	If Not rsads.eof Then
	Dim adsranNum
	Randomize()  
	adsranNum = int(rsads.recordCount*rnd)+1 
	rsads.absoluteposition=adsranNum
	w ("<a href='?aid=url&amp;id="&rsads("id")&"'>"&noubb(rsads("name"))&"</a>")
	End If
	rsads.close
	Set rsads=Nothing
End Function
'随机广告,定义数目
Function adstrs(adsnum,num)
	Dim rsads
	Set rsads=Server.CreateObject("Adodb.Recordset")
	Randomize
	rsads.open"select top "&num&" id,name from 74hu_gogo where typeID="&adsnum&" order by rnd(-(id+" & rnd() & ")) ",conn,1,1
	While Not rsads.EOF
	w ("<a href='?aid=url&amp;id="&rsads("id")&"'>"&noubb(rsads("name"))&"</a><br/>")
	rsads.MoveNext
	Wend
	rsads.close
	Set rsads=Nothing
End Function
'定义广告
Function adsetkf(adnum)
	Dim rsadset
	Set rsadset=Server.CreateObject("Adodb.Recordset")
	rsadset.open"select "&adnum&" from 74hu_control where ID=1",conn,1,1
	If Not rsadset.eof Then
	adsetkf=rsadset(adnum)
	End If
	rsadset.close
	Set rsadset=nothing
End Function

'最新文章
Function newtitle(num,relid)
	Dim gettest,rs1,a
	If relid<>0 Then
	gettest="where classid="&relid
	End If
	Set rs1=Server.CreateObject("Adodb.Recordset")
	rs1.open"select id,title,classid from 74hu_article "&gettest&" order by id desc",conn,1,1
	If rs1.eof Then 
	w ("还没有文章！<br/>")
	Else
	rs1.Move(0)
	a=1
	Do While ((Not rs1.EOF) And a <=num)
	If hu_getLeft Then
	w "<a href=""?aid=art&amp;id="&rs1("id")&""">"&getLeft(noubb(rs1("title")),titlenums)&"</a><br/>"
	Else
	w "<a href=""?aid=art&amp;id="&rs1("id")&""">"&noubb(rs1("title"))&"</a><br/>"
	End If
	rs1.MoveNext
	a=a+1
	Loop
	End If
	rs1.close
	Set rs1=Nothing
End Function
'最热文章
Function hottitle(num,relid)
	dim rs2,b
	If relid<>0 Then
	gettest="where classid="&relid
	End If
	Set rs2 = Server.CreateObject("Adodb.Recordset")
	rs2.open"select id,title,classid from 74hu_article "&gettest&" order by hit desc",conn,1,1
	If rs2.eof Then 
	w ("还没有文章！<br/>")
	Else
	rs2.Move(0)
	b=1
	Do While ((Not rs2.eof) And b <=num)
	If hu_getLeft Then
	w "<a href=""?aid=art&amp;id="&rs2("id")&""">"&getLeft(noubb(rs2("title")),titlenums)&"</a><br/>"
	Else
	w "<a href=""?aid=art&amp;id="&rs2("id")&""">"&noubb(rs2("title"))&"</a><br/>"
	End If
	rs2.MoveNext
	b=b+1
	Loop
	End If
	rs2.close
	Set rs2=Nothing
End Function
'随机文章
Function wendtitle(num,relid)
	Dim rs3
	If relid<>0 Then
	gettest="where classid="&relid
	End If
	Set rs3=Server.CreateObject("Adodb.Recordset")
	Randomize
	rs3.open"select top "&num&" id,title,classid from 74hu_article "&gettest&" order by rnd(-(id*"&rnd()&")) ",conn,1,1
	While Not rs3.eof
	If hu_getLeft Then
	w "<a href=""?aid=art&amp;id="&rs3("id")&""">"&getLeft(noubb(rs3("title")),titlenums)&"</a><br/>"
	Else
	w "<a href=""?aid=art&amp;id="&rs3("id")&""">"&noubb(rs3("title"))&"</a><br/>"
	End If
	rs3.MoveNext
	Wend
	rs3.close
	Set rs3=Nothing
End Function
'翻页菜单2.0
Function turnpage2(aid,add)
	turnpage2="<form name=""f"&Time_r&""" action=""?"" method=""get""><input name=""page"" type=""text"" size=""3"" maxlength=""2""/>"&_
	"<input name=""aid"" type=""hidden"" value="""&aid&"""/>"&add&"<input type=""submit"" value=""跳转""></form>"
End Function
'显示内容
Sub w(str)
	Response.Write str
End Sub
'显示内容且停止输出
Sub wn(str)
	Response.Write str
	Response.End
End Sub
'网页跳转
Sub r(str)
	Response.Redirect str
End Sub
'得到链接
Sub tourl(str,name)
	w getUrl(str,name,"")
End Sub
'得到图片
Sub toimg(str,name)
	w getImg(str,name,"")
End Sub
'改写left 中英文长度取定长修整
Function getLeft(str,len)
	getLeft=hu_title(str,len)
End Function
'获取数据
Function getD(str,def)
	Dim tmp
	tmp=getData(str)
	If hu_isNull(tmp) Then getD=def:Exit Function
	tmp=hu_common(tmp)
	tmp=hu_encode(tmp)
	getD=tmp
End Function
'不过滤获取数据
Function getDD(str,def)
	Dim tmp
	tmp=getData(str)
	If hu_isNull(tmp) Then getDD=def:Exit Function
	getDD=tmp
End Function
'完全过滤获取数据
Function getFilter(str,def)
	Dim tmp
	tmp=getData(str)
	If hu_isNull(tmp) Then getFilter=def:Exit Function
	getFilter=hu_filter(tmp)
End Function
'完全过滤数据
Function setFilter(str)
	setFilter=hu_filter(str)
End Function
'获取数字
Function getN(str,def)
	Dim tmp
	tmp=getData(str)
	If hu_isNull(tmp) Then getN=def:Exit Function
	If Not IsNumeric(tmp) Then getN=def:Exit Function
	getN=int(tmp)'避免非十进制 用clng会溢出
End Function
'从终端获取数据
Function getData(str)
	Dim tmp
	tmp=Trim(Request.QueryString(str))
	If hu_isNull(tmp) Then tmp=Trim(Request.Form(str))
	getData=tmp
End Function
'标题和不使用ubb的内容noubb,后台编辑
Function noubb(str)
	If hu_isNull(str) Then Exit Function
	str=Trim(str)
	str=hu_forShow(str)
	str=changeWord(str)
	str=Replace(str,""," ")
	str=Replace(str,"&nbsp;"," ")
	noubb=str
End Function
'用于链接
Function noubburl(str)
	If hu_isNull(str) Then Exit Function
	str=Trim(str)
	str=hu_decode(str)
	str=changeWord(str)
	str=Replace(str,"&amp;","&")'勿删
	str=Replace(str,"&amp;","&")
	str=Replace(str,"<","")
	str=Replace(str,">","")
	str=Replace(str,"'","")
	str=Replace(str,"""","")
	str=Replace(str,"","")
	str=Replace(str,"&nbsp;","")
	str=Replace(str,"&#35;","#")
	str=Replace(str,"&#58;",":")
	str=Replace(str,"&#61;","=")
	str=Replace(str,"&#63;","?")
	noubburl=str
End Function
'ubb展示
Function ubbcode(str)
	If hu_isNull(str) Then Exit Function
	Dim newstr,re
	newstr=Now
	str=Trim(str)
	str=hu_forShow(str)
	str=changeWord(str)
	str=Replace(str,"&nbsp;"," ")
	str=Replace(str,"[br]","<br/>")
	str=Replace(str,"\\","<br/>")
	str=Replace(str,"[date]",Date)
	str=Replace(str,"[time]",Time)
	str=Replace(str,"[now]",newstr)
	str=Replace(str,"[week]",WeekDayName(DatePart("w",newstr)))'星期几
	str=Replace(str,"[month]",Month(newstr))
	str=Replace(str,"[day]",Day(newstr))
	str=Replace(str,"[hello]",gethello)
	str=Replace(str,"[favor]",getfavor)
	str=Replace(str,"[wapname]",waptitle)
	str=Replace(str,"[wapurl]",wapurl)
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	re.pattern="(\[img\])(.[^\[]*)(\[\/img\])"
	str=re.Replace(str,"<img src=""$2"" alt='.'/>")
	re.pattern="(\[img=(.[^\]]*)\])(.[^\[]*)(\[\/img\])"
	str=re.Replace(str,"<a href=""$3""><img src=""$2"" alt="".""/></a>")
	re.pattern="(\[u\])(.[^\[]*)(\[\/u\])"
	str=re.Replace(str,"<u>$2</u>")
	re.pattern="(\[i\])(.[^\[]*)(\[\/i\])"
	str=re.Replace(str,"<i>$2</i>")
	re.pattern="(\[b\])(.[^\[]*)(\[\/b\])"
	str=re.Replace(str,"<b>$2</b>")
	re.pattern="(\[day=(.[^\]]*)\])(.[^\[]*)(\[\/day\])"
	str=re.Replace(str,getDiff("$2","$3"))
	re.pattern="(\[url\])(.[^\[]*)(\[\/url\])"
	str=re.Replace(str,"<a href=""$2"" >$2</a>")
	re.pattern="(\[url=(.[^\]]*)\])(.[^\[]*)(\[\/url\])"
	str=re.Replace(str,"<a href=""$2"" >$3</a>")
	re.Pattern="(\[m1\])(.[^\[]*)(\[\/m1\])"
	str=re.Replace(str,"<marquee>$2</marquee>")
	re.Pattern="(\[m2\])(.[^\[]*)(\[\/m2\])"
	str=re.Replace(str,"<marquee behavior=""alternate"">$2</marquee>")
	set re=Nothing
	ubbcode=str
End Function
'简单问候语
Function getHello()
	Dim newtime
	newtime=Time
	If newtime < #06:00:00# And newtime >= #00:30:00# Then 
		getHello="凌晨好！"
	ElseIf newtime < #09:00:00# And newtime >= #06:00:00# Then 
		getHello="早上好！"
	ElseIf newtime < #11:30:00# And newtime >= #09:00:00# Then 
		getHello="上午好！"
	ElseIf newtime < #12:30:00# And newtime >= #11:30:00# Then 
		getHello="中午好！"    
	ElseIf newtime < #18:00:00# And newtime >= #12:30:00# Then
		getHello="下午好！"
	ElseIf newtime < #20:00:00# And newtime >= #18:00:00# Then 
		getHello="傍晚好！"      
	ElseIf newtime < #23:30:00# And newtime >= #20:00:00# Then 
		getHello="晚上好！"    
	Else 
		getHello="午夜好！"
	End If 
End Function
'完整问候语
Function getfavor()
	Dim newtime,newmon,newday
	newtime = Time:newmon = month(now):newday = day(now)
	If newtime < #06:00:00# And newtime >= #04:00:00# Then
		getfavor=""&newmon&"月"&newday&"日"&" "&"凌晨好！"
	ElseIf newtime < #09:00:00# And newtime >= #06:00:00# Then
		getfavor=""&newmon&"月"&newday&"日"&" "&"早上好！"
	ElseIf newtime < #11:30:00# And newtime >= #09:00:00# Then
		getfavor=""&newmon&"月"&newday&"日"&" "&"上午好！"
	ElseIf newtime < #12:30:00# And newtime >= #11:30:00# Then
		getfavor=""&newmon&"月"&newday&"日"&" "&"午饭时间到啦。"
	ElseIf newtime < #18:00:00# And newtime >= #12:30:00# Then
		getfavor=""&newmon&"月"&newday&"日"&" "&"下午好！"
	ElseIf newtime < #19:30:00# And newtime >= #18:00:00# Then 
		getfavor=""&newmon&"月"&newday&"日"&" "&"晚饭时间到啦。"
	ElseIf newtime < #23:30:00# And newtime >= #19:30:00# Then
		getfavor=""&newmon&"月"&newday&"日"&" "&"晚上好！"
	Else
		getfavor=""&newmon&"月"&newday&"日"&" "&"夜深注意休息。"
	End If
End Function
'统一时间 2008.8.8 20:08
Function fordate(str)
	fordate=hu_dateFormat(str,1)
End Function
'统一时间 8-8 20:08
Function fordate2(str)
	fordate2=hu_dateFormat(str,2)
End Function
'网页头部
Sub getHead(str, ver)
	Select Case ver
	Case 1
		Response.ContentType = "text/vnd.wap.wml; charset=utf-8"
		w "<?xml version=""1.0"" encoding=""utf-8""?>" &_
			"<!DOCTYPE wml PUBLIC ""-//WAPFORUM//DTD WML 1.1//EN"" ""http://www.wapforum.org/DTD/wml_1.1.xml"">" &_
			"<wml><head>" & str
	Case 2
		Response.ContentType = "text/html; charset=utf-8"
		w "<?xml version=""1.0"" encoding=""utf-8""?>" &_
			"<!DOCTYPE html PUBLIC ""-//WAPFORUM//DTD XHTML Mobile 1.0//EN"" ""http://www.wapforum.org/DTD/xhtml-mobile10.dtd"">" &_
			"<html xmlns=""http://www.w3.org/1999/xhtml""><head>" &_
			"<meta http-equiv=""Content-Type"" content=""text/html"" charset=""utf-8""/>" & str
	Case 0
		Response.ContentType = "text/html; charset=utf-8"
		w "<?xml version=""1.0"" encoding=""utf-8""?>" &_
			"<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" &_
			"<html xmlns=""http://www.w3.org/1999/xhtml""><head>" &_
			"<meta http-equiv=""Content-Type"" content=""text/html"" charset=""utf-8""/>"& str
	End Select
End Sub
'网页尾部
Sub getEnd(str, ver)
	Select Case ver
	Case 1
		w str & "</p></card></wml>"
	Case 2,0
		w str & "</body></html>"
	End Select
	Response.End
End Sub
'关闭数据库连接
Sub getClose()
	conn.close
	set conn=nothing
End Sub
'网页标题
Sub getTitle(str, ver)
	Select Case ver
	Case 1
		w "<card title=""" & str & """>"
	Case 2,0
		w "<title>" & str & "</title>"
	End Select
End Sub
'构造链接
Function getUrl(str, name, ex)
	Dim newstr
	newstr = ""
	If Not hu_isNull(ex) Then newstr = ex
	getUrl = "<a href=""" & str & """ " & newstr & ">" & name & "</a>"
End Function
'构造图片
Function getImg(str, name, ex)
	Dim newstr
	newstr = ""
	If Not hu_isNull(ex) Then newstr = ex
	getImg = "<img src=""" & str & """ title=""" & name & """ alt=""loading.."" " & ex & " />"
End Function
'清除缓存
Sub cache(str)
	If str Then Exit Sub
	Response.Buffer = True
	Response.Expires = 0
	Response.ExpiresAbsolute = Now() - 1
	Response.CacheControl = "no-cache"
	Response.AddHeader "Expires",Date()
	Response.AddHeader "Pragma","no-cache"
	Response.AddHeader "Cache-Control","private, no-cache, must-revalidate"
End Sub
'时间比较
Function getDiff(day,str)
	'以后可以精细到秒
	day = Trim(day)
	If Not isDate(day) Then Exit Function
	Dim newstr
	newstr = DateDiff("d",date,Cdate(day))
	getDiff = "距" & str & "还有" & newstr & "天"
End Function
'敏感词过滤
Function changeWord(str)
	changeWord = hu_changeWord(str,hu_badWord,"[滤]")
End Function
'IP封锁
Sub ipLock(str)
	Dim IpArray,WhyIpLock,IpSQL,IpRS
	IpArray=split(str,".")
	IpSQL="SELECT iplock From 74hu_IpLock Where  "& _
	" (ipsame=4 and ip1="&Cint(IpArray(0))&" and ip2="&Cint(IpArray(1))&" and ip3="&Cint(IpArray(2))&" and ip4="&Cint(IpArray(3))&" )  "& _
	" Or (ipsame=3 and  ip1="&Cint(IpArray(0))&"  and  ip2="&Cint(IpArray(1))&"  and  ip3="&Cint(IpArray(2))&" )   "& _
	" Or (ipsame=2 and ip1="&Cint(IpArray(0))&" and ip2="&Cint(IpArray(1))&" ) Or (ipsame=1 and ip1="&Cint(IpArray(0))&" ) Order By ipid "
	Set IpRS=Conn.execute(IpSQL)
	If Not (IpRS.bof or IpRS.eof) Then
		WhyIpLock=split(IpRS("iplock"),"|")
		Response.write "<card title=""出错了""><p>你使用的IP段或IP地址已被封锁<br/>封锁原因:"&WhyIpLock(1)&"<br/>封锁时间:"&WhyIpLock(0)&"</p></card></wml>"
		Response.End
		Set Conn=nothing
	End If
	Set IpRS=Nothing
End Sub
'流量统计
Sub setStatistics(str)
	Dim HU_users,HU_userip,rsip
	HU_users="七色虎"
	HU_userip=str
	Set rsip = Server.CreateObject("ADODB.Recordset")
	rsip.open"select HU_Date,HU_Tod,HU_Today from 74hu_counter",conn,1,1
	HU_Date=rsip("HU_Date")
	if HU_Date<>date() then
		HU_day=date()-1
		conn.Execute"Update 74hu_counter set HU_Today=0,HU_Browser=0,HU_Date='"&date()&"',HU_Yays=HU_Yays+1,HU_Yesterday="&rsip("HU_Today")&""
		conn.Execute"delete from 74hu_iprr"
	else
		conn.Execute"Update 74hu_counter set HU_Browser=HU_Browser+1"
		if conn.execute("select HU_userip from 74hu_iprr where HU_userip='"&HU_userip&"'").eof then
			conn.Execute"insert into 74hu_iprr(HU_Userip,Users) values('"&HU_userip&"','"&HU_users&"')"
			conn.Execute"Update 74hu_counter set HU_counter=HU_counter+1,HU_Today=HU_Today+1"
		end if
	end if
	conn.Execute"Update 74hu_counter set HU_Tod="&rsip("HU_Today")&" where "&rsip("HU_Tod")&"<"&rsip("HU_Today")
	conn.Execute"Update 74hu_counter set HU_Browsers=HU_Browsers+1"
	rsip.close
	set rsip=nothing
End Sub
'2.0编辑后写入数据库
Function forSaveByWeb(str)
	If hu_isNull(str) Then Exit Function
	str=Trim(str)
	str=Replace(str,"&nbsp;"," ")
	str=Replace(str,"&amp;","&#38;")
	str=Replace(str,"$$","$")'兼容1.0
	str=Replace(str,"","")
	str=Replace(str,vbnewline,"\\")
	str=Replace(str,VbCrLf,"\\")
	forSaveByWeb=str
End Function

' 
' 底层函数，外部不要直接引用
' 要求：低耦合
' 函数命名：hu_myNameIsHu

'改写IsNull
Function hu_isNull(str)
	hu_isNull = False
	Select Case VarType(str)
	Case vbEmpty, vbNull
		hu_isNull = True : Exit Function
	Case vbstring
		If str="" Then hu_isNull = True : Exit Function
	Case vbObject
		If TypeName(str)="Nothing" Or TypeName(str)="Empty" Then hu_isNull = True : Exit Function
	Case vbArray,8194,8204,8209
		If Ubound(str)=-1 Then hu_isNull = True : Exit Function
	End Select
End Function
'格式化时间
Function hu_dateFormat(str, style)
	If Not IsDate(str) Then Exit Function
	Select Case style
	Case 1'2008.8.8 20:08
		hu_dateFormat = year(str) & "." & month(str) & "." & day(str) & " "
		If Hour(str) < 10 Then hu_dateFormat = hu_dateFormat&"0"
		hu_dateFormat = hu_dateFormat&Hour(str)&":"
		If Minute(str) < 10 Then hu_dateFormat = hu_dateFormat&"0"
		hu_dateFormat = hu_dateFormat&Minute(str)
	Case 2'8-8 20:08
		hu_dateFormat = hu_dateFormat&month(str) & "-"
		hu_dateFormat = hu_dateFormat&day(str)&" "
		hu_dateFormat = hu_dateFormat&Hour(str)&":"
		If Minute(str) < 10 Then hu_dateFormat = hu_dateFormat&"0"
		hu_dateFormat = hu_dateFormat&Minute(str)	
	End Select
End Function
'字符过滤函数
Function hu_changeWord(str, badStandard, changedWord)
	Dim s,i
	s=split(badStandard,",")
	For i=Lbound(s) To ubound(s)
	str=replace(str,s(i),changedWord)
	Next
	hu_changeWord=str
End Function
'替换bug字符 - 展现
Function hu_forShow(str)
	If hu_isNull(str) Then Exit Function
	str=hu_decode(str)
	If hu_style Then hu_forShow=str:Exit Function
	str=Replace(str,"&#","_74_asp_")
	str=Replace(str,"&amp;","_74_amp_")
	str=Replace(str,"&","&amp;")
	str=Replace(str,"_74_amp_","&amp;")
	str=Replace(str,"$$","_74_my_")
	str=Replace(str,"$","$$")
	str=Replace(str,"_74_my_","$$")
	str=Replace(str,"<","&lt;")
	str=Replace(str,">","&gt;")
	str=Replace(str,"'","&apos;")
	str=Replace(str,"""","&quot;")
	str=Replace(str,"_74_asp_","&#")
	hu_forShow=str
End Function
'符号写入数据库
Function hu_common(str)
	If hu_isNull(str) Then Exit Function
	str=Replace(str,"&","_74_aaa_")
	str=Replace(str,"#","&#35;")
	str=Replace(str,"	","&#9;")
	str=Replace(str," ","&#32;")
	str=Replace(str,"'","&#39;")
	str=Replace(str,"""","&#34;")
	str=Replace(str,"%","&#37;")
	str=Replace(str,"*","&#42;")
	str=Replace(str,":","&#58;")
	str=Replace(str,"<","&#60;")
	str=Replace(str,"=","&#61;")
	str=Replace(str,">","&#62;")
	str=Replace(str,"?","&#63;")
	str=Replace(str,vbnewline,"&#13;&#10;")
	str=Replace(str,VbCrLf,"&#13;&#10;")
	str=Replace(str,"_74_aaa_","&#38;")
	hu_common=str
End Function
'过滤用户数据
Function hu_encode(str)
	If hu_isNull(str) Then Exit Function
	str=Replace(str,"74hu_","74_hu_",1,-1,1)'数据库表前缀替换
	str=Replace(str,"and","_74_an_",1,-1,1)
	str=Replace(str,"or","_74_or_",1,-1,1)
	str=Replace(str,"from","_74_fr_",1,-1,1)
	str=Replace(str,"mid","_74_mi_",1,-1,1)
	str=Replace(str,"update","_74_up_",1,-1,1)
	str=Replace(str,"exec","_74_ex_",1,-1,1)
	str=Replace(str,"select","_74_se_",1,-1,1)
	str=Replace(str,"insert","_74_in_",1,-1,1)
	str=Replace(str,"delete","_74_de_",1,-1,1)
	str=Replace(str,"drop","_74_dr_",1,-1,1)
	str=Replace(str,"create","_74_cr_",1,-1,1)
	str=Replace(str,"eval","_74_ev_",1,-1,1)
	str=Replace(str,"command","_74_co_",1,-1,1)
	str=Replace(str,"dir","_74_di_",1,-1,1)
	str=Replace(str,"truncate","_74_tr_",1,-1,1)
	str=Replace(str,"xp_","_74_xp_",1,-1,1)
	str=Replace(str,"sp_","_74_sp_",1,-1,1)
	str=Replace(str,"master","_74_ma_",1,-1,1)
	str=Replace(str,"declare","_74_dec_",1,-1,1)
	str=Replace(str,"count","_74_cou_",1,-1,1)
	str=Replace(str,"char","_74_ch_",1,-1,1)
	str=Replace(str,"unicode","_74_un_",1,-1,1)
	str=Replace(str,"ascii","_74_as_",1,-1,1)
	str=Replace(str,"cmd","_74_cm_",1,-1,1)
	str=Replace(str,"法轮","[滤]")'国内服务器拒绝写入
	hu_encode=str
End Function
'还原用户数据
Function hu_decode(str)
	If hu_isNull(str) Then Exit Function
	str=Replace(str,"_74_hu_","74hu_")
	str=Replace(str,"_74_an_","and")
	str=Replace(str,"_74_or_","or")
	str=Replace(str,"_74_fr_","from")
	str=Replace(str,"_74_mi_","mid")
	str=Replace(str,"_74_up_","update")
	str=Replace(str,"_74_ex_","exec")
	str=Replace(str,"_74_se_","select")
	str=Replace(str,"_74_in_","insert")
	str=Replace(str,"_74_de_","delete")
	str=Replace(str,"_74_dr_","drop")
	str=Replace(str,"_74_cr_","create")
	str=Replace(str,"_74_ev_","eval")
	str=Replace(str,"_74_co_","command")
	str=Replace(str,"_74_di_","dir")
	str=Replace(str,"_74_tr_","truncate")
	str=Replace(str,"_74_xp_","xp_")
	str=Replace(str,"_74_sp_","sp_")
	str=Replace(str,"_74_ma_","master")
	str=Replace(str,"_74_dec_","declare")
	str=Replace(str,"_74_cou_","count")
	str=Replace(str,"_74_ch_","char")
	str=Replace(str,"_74_un_","unicode")
	str=Replace(str,"_74_as_","ascii")
	str=Replace(str,"_74_cm_","cmd")
	hu_decode=str
End Function
'用于搜索,登陆过滤等
Function hu_filter(str)
	If hu_isNull(str) Then Exit Function
	str=Replace(str,"'","",1,-1,1)
	str=Replace(str,"""","",1,-1,1)
	str=Replace(str,":","",1,-1,1)
	str=Replace(str,"*","",1,-1,1)
	str=Replace(str,"<","",1,-1,1)
	str=Replace(str,">","",1,-1,1)
	str=Replace(str,"or","",1,-1,1)
	str=Replace(str,"74hu_","",1,-1,1)
	str=Replace(str,"and","",1,-1,1)
	str=Replace(str,"from","",1,-1,1)
	str=Replace(str,"mid","",1,-1,1)
	str=Replace(str,"update","",1,-1,1)
	str=Replace(str,"exec","",1,-1,1)
	str=Replace(str,"select","",1,-1,1)
	str=Replace(str,"insert","",1,-1,1)
	str=Replace(str,"delete","",1,-1,1)
	str=Replace(str,"drop","",1,-1,1)
	str=Replace(str,"create","",1,-1,1)
	str=Replace(str,"eval","",1,-1,1)
	str=Replace(str,"command","",1,-1,1)
	str=Replace(str,"dir","",1,-1,1)
	str=Replace(str,"truncate","",1,-1,1)
	str=Replace(str,"xp_","",1,-1,1)
	str=Replace(str,"sp_","",1,-1,1)
	str=Replace(str,"master","",1,-1,1)
	str=Replace(str,"declare","",1,-1,1)
	str=Replace(str,"count","",1,-1,1)
	str=Replace(str,"char","",1,-1,1)
	str=Replace(str,"unicode","",1,-1,1)
	str=Replace(str,"ascii","",1,-1,1)
	str=Replace(str,"cmd","",1,-1,1)
	str=Replace(str,"法轮","")'国内服务器拒绝写入
	hu_filter=str
End Function
'改写left
Function hu_title(str, strlen)
	If hu_isNull(str) Then hu_title = "":Exit Function
	Dim l, t, c, i, strTemp
	str = Replace(Replace(Replace(Replace(Replace(str,"&nbsp;"," "),"&quot;",Chr(34)),"&gt;",">"),"&lt;","<"),"&apos;","'")
	l = Len(str):t = 0:strTemp = str:strlen = CLng(strlen)
	For i = 1 To l:c = Abs(Asc(Mid(str, i, 1)))
		If c = 1 Then:t = t + 1:Else:t = t + 0.6:End If'这里的0.6可酌情修改，考虑字符占位不同
		If t >= strlen Then:strTemp = Left(str, i):Exit For:End If
	Next:If strTemp <> str Then:strTemp = strTemp & "..":End If
	hu_title = Replace(Replace(Replace(Replace(Replace(strTemp," ","&nbsp;"),Chr(34),""""),">","&gt;"),"<","&lt;"),"'","&apos;")
End Function
%>
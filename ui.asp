<%
'
'	七色虎建站系统
'	表现层文件Db.asp
'	用于展现网站
'	v1.2.4.143a
'	2011.9.3

' 
' 表现层函数，外部可以直接引用
' 要求：共同属性写入底层，这部分只是用于展现
' 函数命名：showMyName

'Wap1.0首页
Sub showIndex()
	w "<card title="""&waptitle&"""><p align="""&wapconst&""">"
	if wapfavor="1" then w ""&getfavor&"<br/>"
	if len(countdown) > 7 then w getDiff(countdown,countname)&"<br/>"
	If Len(waplogo) > 7 Then w "<img src="""&waplogo&""" alt="""&waptitle&"""/><br/>"
	if wapgonggao="1" then w "<a href=""?aid=gonggao""><img src=""images/msg.gif"" alt="".""/>网站发布最新公告!</a><br/>"
	Dim rs,j
	Set rs = Server.CreateObject("adodb.recordset")
	rs.open "select lx,class,wmltxt,relid,br,num,classid from 74hu_class where parent=0 order by pid asc", conn, 1, 1
	If rs.eOF Then
		w "网站建设中..<br/>"
	else
		rs.Move (0)
		j = 1
		Do While Not rs.eOF
			Select Case rs("lx")
			Case 2 w ubbcode(rs("wmltxt"))
			Case 9 w rs("wmltxt")
			Case 8 Call adstr(1)
			Case 10 Call newtitle(rs("num"), rs("relid"))
			Case 11 Call hottitle(rs("num"), rs("relid"))
			Case 12 Call wendtitle(rs("num"), rs("relid"))
			Case 0 w "<a href=""?aid=class&amp;id="&rs("classid")&""">"&noubb(rs("class"))&"</a>"
			Case 1 w "<a href=""?aid=list&amp;id="&rs("relid")&""">"&noubb(rs("class"))&"</a>"
			Case 19 w "<input emptyok=""true"" name=""keyword"" value="""" title=""请输入关键词""/><br/>"&_
				"搜<anchor>文章<go href=""search.asp"" method=""post""><postfield name=""keyword"" value=""$(keyword)""/>"&_
				"<postfield name=""sear"" value=""0""/></go></anchor>"&_
				"搜<anchor>网页<go href=""http://u.yicha.cn/union/x.jsp"" method=""post""><postfield name=""keyword"" value=""$(keyword)""/>"&_
				"<postfield name=""site"" value=""2145930044""/><postfield name=""p"" value=""p""/></go></anchor>"
			End Select
			If rs("br") = "1" Then w "<br/>"
			rs.MoveNext
			j = j + 1
		Loop
	end If
	rs.close
	Set rs = nothing
	If waplink = 1 Then
		Dim Rslc,aaa
		Set Rslc = Server.CreateObject("ADODB.Recordset")
		Sqlink = "select ID,namt from 74hu_link Where active =0 order by HU_time desc"
		Rslc.open Sqlink, conn, 1, 1
		If Rslc.EOF Then w "暂无首链！<br/>"
		aaa = 1
		Do While ((Not Rslc.EOF) And aaa <= 8)
			w "<a href=""?aid=link&amp;id=" & Rslc("id") & "&amp;act=view"">" & noubb(Rslc("NAMT")) & "</a> "
			If aaa Mod 4 = 0 And aaa <> Rslc.RecordCount Then w "<br/>"
			Rslc.MoveNext
			aaa = aaa + 1
		Loop
		Rslc.Close
		Set Rslc = Nothing
	End If
End Sub
'Wap2.0首页
Sub showsIndex()
	w "<title>"&waptitle&"</title></head><body><div class=""main"">"
	if wapfavor="1" then w getfavor&"<br/>"
	if len(countdown) > 7 then w "<font color=""red"">"&getDiff(countdown,countname)&"</font><br/>"
	If Len(waplogo) > 7 Then w "<img src='"&waplogo&"' alt='"&waptitle&"'/><br/>"
	if wapgonggao="1" then w "<a href='?aid=gonggao'><img src='images/msg.gif' alt='.'/>网站发布最新公告!</a><br/>"
	Dim rs,j
	Set rs = Server.CreateObject("adodb.recordset")
	rs.open "select lx,class,wmltxt,relid,br,num,classid from 74hu_class where parent=0 order by pid asc", conn, 1, 1
	If rs.eOF Then
	w "网站建设中..<br/>"
	else
	rs.Move (0)
	j = 1
	Do While Not rs.eOF
		Select Case rs("lx")
		Case 2 w ubbcode(rs("wmltxt"))
		Case 9 w rs("wmltxt")
		Case 8 Call adstr(1)
		Case 10 Call newtitle(rs("num"), rs("relid"))
		Case 11 Call hottitle(rs("num"), rs("relid"))
		Case 12 Call wendtitle(rs("num"), rs("relid"))
		Case 0 w "<a href=""?aid=class&amp;id="&rs("classid")&""">"&noubb(rs("class"))&"</a>"
		Case 1 w "<a href=""?aid=list&amp;id="&rs("relid")&""">"&noubb(rs("class"))&"</a>"
		Case 19 w "<form id=""sch"" method=""post"" action=""search.asp""><input emptyok=""true"" name=""keyword"" value="""" title=""请输入关键词""/><br/>"&_
			"<input type=""submit"" value=""搜文章""/>&nbsp;<a href=""search.asp"">更多搜索</a></form>"
		End Select
	If rs("br") = "1" Then w "<br/>"
	rs.MoveNext
	j = j + 1
	Loop
	end If
	rs.close
	Set rs = nothing
	If waplink = 1 Then
	Dim Rslc,aaa
	Set Rslc = Server.CreateObject("ADODB.Recordset")
	Sqlink = "select ID,namt from 74hu_link Where active =0 order by HU_time desc"
	Rslc.open Sqlink, conn, 1, 1
	If Rslc.EOF Then w "暂无首链！<br/>"
	aaa = 1
	Do While ((Not Rslc.EOF) And aaa <= 8)
	w "<a href=""?aid=link&amp;id=" & Rslc("id") & "&amp;act=view"">" & noubb(Rslc("NAMT")) & "</a>&nbsp;"
	If aaa Mod 4 = 0 And aaa <> Rslc.RecordCount Then w "<br/>"
	Rslc.MoveNext
	aaa = aaa + 1
	Loop
	Rslc.Close
	Set Rslc = Nothing
	End If
End Sub
'Wap1.0文章页
Sub showArticle()
	Dim p,rs,sql
	p=getN("p",1)
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="Select title,test,hit,smspin,classid,HU_author,HU_date from 74hu_article where id="&id
	rs.open sql,conn,1,3
	if rs.eof then
	rs.close
	set rs=Nothing
	r "?aid=index"
	end if
	Dim ids,rss
	ids=rs("classid")
	Set rss = Server.CreateObject("ADODB.Recordset")
	sql="Select class from 74hu_list where classid="&ids
	rss.open sql,conn,3,1 
	if rss.eof then
	rss.close
	set rss=Nothing
	r "?aid=index"
	end if
	rs("hit")=rs("hit")+1
	rs.update()
	w "<card title='"&noubb(rs("title"))&"-"&rss("class")&"'><p>"
	if wapgonggao="1" then w "<a href='?aid=gonggao'><img src='images/msg.gif' alt='.'/>网站发布最新公告!</a><br/>"
	w "-<a href='?aid=index'>首页</a>-<a href='?aid=list&amp;id="&ids&"&amp;page="&p&"'>"&rss("class")&"</a>-正文<br/>"
	Dim Counts,Content,pageWordNum,StartWord,Length,PageAll,page,i,ccc,sss
	Counts=rs("smspin")
	w ""&noubb(rs("title"))&"<br/>-----------<br/>"
	if adsetkf("ads1")=1 then
	call adstr(1)
	w "<br/>"
	end if
	w "内容来源:"&rs("HU_author")&"<br/>["&fordate(rs("HU_date"))&"]<br/>"
	Content=rs("test")
	pageWordNum=viewtnums
	StartWord = 1
	Length=len(Content)
	PageAll=(Length+PageWordNum-1)\PageWordNum
	page=getN("page",1)
	if page<1 then page=1
	i=int(page-1)
	if page>PageAll then page=PageAll
	if isnull(i) or IsNumeric(i)=False then i=0
	ccc=instr(content,"||")
	if ccc>0 then
	sss=split(content,"||")
	PageAll=ubound(sss)+1
	if i>PageAll-1 then i=PageAll-1
	content = sss(i)
	else
	if clng(i)>int(PageAll) then i=PageAll-1
	Content = mid(Content,StartWord+i*PageWordNum,PageWordNum)
	end if
	w  ubbcode(content)
	if 0<=i<PageAll then w "<br/>"
	if cint(i)<cint(PageAll)-1 then w "<a href='?aid=art&amp;id=" & id & "&amp;page=" & i+2 & "&amp;p=" & p & "'>下页</a>"&"&nbsp;" 
	if cint(i)>0 then w "<a href='?aid=art&amp;id=" & id & "&amp;page=" & i & "&amp;p=" & p & "'>上页</a>"
	if PageAll>1 then w "(" & i+1 & "/" & PageAll & ")"&_
		"<br/>第<input name=""i"" type=""text"" format=""*N"" emptyok=""true"" size=""2"" value="""" maxlength=""2""/>页"&_
		"<anchor>跳转<go href=""?aid=art&amp;id="&id&"&amp;p="&p&""" accept-charset=""utf-8"">"&_
		"<postfield name=""page"" value=""$(i)""/></go></anchor><br/>"

	w "-----------<br/>※快速评论：<br/>"&_
		"<input type=""text"" name=""pl"&Time_r&""" title=""输入内容"" value="""" maxlength=""200""/><br/>"&_
		"<anchor title=""确定"">提交<go method=""get"" href=""?aid=diss&amp;id="&id&"&amp;p="&p&""">"&_
		"<postfield name=""pl"" value=""$(pl"&Time_r&")""/>"&_
		"</go></anchor>"&_
		"<a href=""?aid=dis&amp;id="&id&"&amp;p="&p&""">网友评论("&Counts&")条</a><br/>"
	Dim rs1,sql2,rs2
	set rs1=server.createobject("adodb.recordset")
	sql2="select top 1 id,test,title from 74hu_article where classid="&ids&" and id<"&id&" order by id desc"
	rs1.open sql2,conn,3,1
	if rs1.recordcount>0 then
	w "<a href=""?aid=art&amp;id="&rs1("id")&"&amp;p="&p&""">&gt;&gt;"&noubb(rs1("title"))&"</a><br/>"
	end if
	rs1.close
	set rs1=nothing
	set rs2=server.createobject("adodb.recordset")
	sql2="select top 1 id,test,title from 74hu_article where classid="&ids&" and id>"&id&" order by id asc"
	rs2.open sql2,conn,3,1
	if rs2.recordcount>0 then
	w "<a href=""?aid=art&amp;id="&rs2("id")&"&amp;p="&p&""">&lt;&lt;"&noubb(rs2("title"))&"</a><br/>"
	end if
	w "[相关内容]<br/>"
	call wendtitle(3,ids)
	if adsetkf("ads2")=1 then
	call adstr(2)
	w "<br/>"
	end if
	w wapurl&" ["&fordate2(Now)&"]"
	rs.close
	set rs=nothing
	rss.close
	set rss=nothing
	rs2.close
	set rs2=nothing
End Sub
'Wap2.0文章页
Sub showsArticle()
	Dim p,rs,sql
	p=getN("p",1)
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="Select title,test,hit,smspin,classid,HU_author,HU_date from 74hu_article where id="&id
	rs.open sql,conn,1,3
	if rs.eof then
	rs.close
	set rs=Nothing
	r "?aid=index"
	end if
	Dim ids,rss
	ids=rs("classid")
	Set rss = Server.CreateObject("ADODB.Recordset")
	sql="Select class from 74hu_list where classid="&ids
	rss.open sql,conn,3,1 
	if rss.eof then
	rss.close
	set rss=Nothing
	r "?aid=index"
	end if
	rs("hit")=rs("hit")+1
	rs.update()
	w "<title>"&noubb(rs("title"))&"-"&rss("class")&"</title></head><body><div class=""main"">"
	if wapgonggao="1" then w "<a href='?aid=gonggao'><img src='images/msg.gif' alt='.'/>网站发布最新公告!</a><br/>"
	w "<div class=""nav"">-<a href='?aid=index'>首页</a>-<a href='?aid=list&amp;id="&ids&"&amp;page="&p&"'>"&rss("class")&"</a>-正文</div>"
	Dim Counts,Content,pageWordNum,StartWord,Length,PageAll,page,i,ccc,sss
	Counts=rs("smspin")
	w "<div class=""tle"">"&noubb(rs("title"))&"</div>-----------<br/>"
	if adsetkf("ads1")=1 then
	call adstr(1)
	w "<br/>"
	end if
	w "内容来源:"&rs("HU_author")&"<br/>["&rs("HU_date")&"]<br/>"
	Content=rs("test")
	pageWordNum=viewtnums
	StartWord = 1
	Length=len(Content)
	PageAll=(Length+PageWordNum-1)\PageWordNum
	page=getN("page",1)
	if page<1 then page=1
	i=int(page-1)
	if page>PageAll then page=PageAll
	if isnull(i) or IsNumeric(i)=False then i=0
	ccc=instr(content,"||")
	if ccc>0 then
	sss=split(content,"||")
	PageAll=ubound(sss)+1
	if i>PageAll-1 then i=PageAll-1
	content = sss(i)
	else
	if clng(i)>int(PageAll) then i=PageAll-1
	Content = mid(Content,StartWord+i*PageWordNum,PageWordNum)
	end if
	w ubbcode(content)
	if 0<=i<PageAll then w "<br/>"
	if cint(i)<cint(PageAll)-1 then w "<a href='?aid=art&amp;id=" & id & "&amp;page=" & i+2 & "&amp;p=" & p & "'>下页</a>"&"&nbsp;" 
	if cint(i)>0 then w "<a href='?aid=art&amp;id=" & id & "&amp;page=" & i & "&amp;p=" & p & "'>上页</a>"
	if PageAll>1 then w "(" & i+1 & "/" & PageAll & ")" &"<br/>"&turnpage2("art","<input type=""hidden"" name=""id"" value="""&id&"""/><input type=""hidden"" name=""p"" value="""&p&"""/>")
	w "<div class=""nav"">※快速评论：</div><form name=""f"&Time_r&""" action=""?"" method=""get""><input type=""text"" name=""pl"" value="""" maxlength=""200""/>"&_
	  "<input type=""hidden"" name=""aid"" value=""diss""/><input type=""hidden"" name=""id"" value="""&id&"""/><input type=""hidden"" name=""p"" value="""&p&"""/>"&_
	  "<br/><input type=""submit"" value=""提交""/></form><a href=""?aid=dis&amp;id="&id&"&amp;p="&p&""">网友评论("&Counts&")条</a><br/>"
	Dim rs1,rs2
	set rs1=server.createobject("adodb.recordset")
	sql="select top 1 id,test,title from 74hu_article where classid="&ids&" and id<"&id&" order by id desc"
	rs1.open sql,conn,3,1
	if rs1.recordcount>0 then
	w "<a href=""?aid=art&amp;id="&rs1("id")&"&amp;p="&p&""">&gt;&gt;"&noubb(rs1("title"))&"</a><br/>"
	end if
	rs1.close
	set rs1=nothing
	set rs2=server.createobject("adodb.recordset")
	sql="select top 1 id,test,title from 74hu_article where classid="&ids&" and id>"&id&" order by id asc"
	rs2.open sql,conn,3,1
	if rs2.recordcount>0 then
	w "<a href=""?aid=art&amp;id="&rs2("id")&"&amp;p="&p&""">&lt;&lt;"&noubb(rs2("title"))&"</a><br/>"
	end if
	w "<div class=""nav"">[相关内容]</div>"
	call wendtitle(3,ids)
	if adsetkf("ads2")=1 then
	call adstr(2)
	w "<br/>"
	end if
	w wapurl&" ["&fordate2(Now)&"]"
	rs.close
	set rs=nothing
	rss.close
	set rss=nothing
	rs2.close
	set rs2=nothing
End Sub
'Wap1.0 新页面页
Sub showClass()
	Dim rs1,sql1
	Set rs1 = Server.CreateObject("ADODB.Recordset")
	sql1="Select class from 74hu_class where classid="&id
	rs1.open sql1,conn,1,1 
	if rs1.eof then r "?aid=index"
	Dim classname,rs,j
	classname=rs1("class")
	rs1.close
	set rs1=nothing
	w "<card title='"&classname&"-"&waptitle&"'><p>"
	set rs = server.createobject("adodb.recordset")
	rs.open"select lx,class,classid,wmltxt,num,relid,br from 74hu_class where parent="&id&" order by pid asc",conn,1,1
	if rs.eof then 
	w "栏目建设中..<br/>"
	else
	rs.Move(0)
	j=1
	do while not rs.EOF 
	Select Case rs("lx")
	Case 2 w ubbcode(rs("wmltxt"))
	Case 9 w rs("wmltxt")
	Case 8 Call adstr(1)
	Case 10 Call newtitle(rs("num"), rs("relid"))
	Case 11 Call hottitle(rs("num"), rs("relid"))
	Case 12 Call wendtitle(rs("num"), rs("relid"))
	Case 0 w "<a href=""?aid=class&amp;id="&rs("classid")&""">"&noubb(rs("class"))&"</a>"
	Case 1 w "<a href=""?aid=list&amp;id="&rs("relid")&""">"&noubb(rs("class"))&"</a>"
	Case 19 w "<input emptyok=""true"" name=""keyword"" value="""" title=""请输入关键词""/><br/>"&_
			"搜<anchor>文章<go href=""search.asp"" method=""post""><postfield name=""keyword"" value=""$(keyword)""/>"&_
			"<postfield name=""sear"" value=""0""/></go></anchor>"&_
			"搜<anchor>网页<go href=""http://u.yicha.cn/union/x.jsp"" method=""post""><postfield name=""keyword"" value=""$(keyword)""/>"&_
			"<postfield name=""site"" value=""2145930044""/><postfield name=""p"" value=""p""/></go></anchor>"
	End Select
	if rs("br")="1" then w "<br/>"
	rs.MoveNext
	j=j+1
	loop
	end if
	rs.close
	set rs=nothing
	w "-----------<br/>"
	if adsetkf("ads3")=1 then call adstrs(3,2)
End Sub
'Wap2.0 新页面页
Sub showsClass()
	Dim rs1,sql1
	Set rs1 = Server.CreateObject("ADODB.Recordset")
	sql1="Select class from 74hu_class where classid="&id
	rs1.open sql1,conn,1,1 
	if rs1.eof then
	r "?aid=index"
	end if
	Dim classname,rs,j
	classname=rs1("class")
	rs1.close
	set rs1=nothing
	w "<title>"&classname&"-"&waptitle&"</title></head><body><div class=""main"">"
	set rs = server.createobject("adodb.recordset")
	rs.open"select lx,class,classid,wmltxt,num,relid,br from 74hu_class where parent="&id&" order by pid asc",conn,1,1
	if rs.eof then 
	w "栏目建设中..<br/>"
	else
	rs.Move(0)
	j=1
	do while not rs.EOF 
	Select Case rs("lx")
	Case 2 w ubbcode(rs("wmltxt"))
	Case 9 w rs("wmltxt")
	Case 8 Call adstr(1)
	Case 10 Call newtitle(rs("num"), rs("relid"))
	Case 11 Call hottitle(rs("num"), rs("relid"))
	Case 12 Call wendtitle(rs("num"), rs("relid"))
	Case 0 w "<a href=""?aid=class&amp;id="&rs("classid")&""">"&noubb(rs("class"))&"</a>"
	Case 1 w "<a href=""?aid=list&amp;id="&rs("relid")&""">"&noubb(rs("class"))&"</a>"
	Case 19 w "<form id=""sch"" method=""post"" action=""search.asp""><input emptyok=""true"" name=""keyword"" value="""" title=""请输入关键词""/><br/>"&_
		"<input type=""submit"" value=""搜文章""/>&nbsp;<a href=""search.asp"">更多搜索</a></form>"
	End Select
	if rs("br")="1" then w"<br/>"
	rs.MoveNext
	j=j+1
	loop
	end if
	rs.close
	set rs=nothing
	w "-----------<br/>"
	if adsetkf("ads3")=1 then call adstrs(3,2)
End Sub
' Wap1.0评论页
Sub showDiscuss()
	Dim p,rs
	p=getN("p",1)
	w "<card title=""网友跟贴""><p>-<a href=""?aid=index"">首页</a>-<a href=""?aid=art&amp;id="&id&"&amp;p="&p&""">查看原文</a>-跟贴"&_
		"<br/>发表评论：<br/>"&_
		"<input type=""text"" name=""pl"&Timer_r&""" title=""输入内容"" value="""" maxlength=""200""/><br/>"&_
		"<anchor title=""确定"">提交"&_
		"<go method=""get"" href=""?aid=diss&amp;id="&id&"&amp;p="&p&""">"&_
		"<postfield name=""pl"" value=""$(pl"&Time_r&")""/>"&_
		"</go></anchor><br/>【网友评论区】<br/>"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open"select id,pl,pltime,ag,da from 74hu_pl where smsid="&id&" order by id desc",conn,3,1
	If Not rs.eof Then
	Dim PageSize,gopage,Count,page,pagecount,i
	PageSize=8
	gopage="?aid=dis&amp;id="&id&"&amp;p="&p&"&amp;"
	Count=conn.execute("Select count(ID) from 74hu_pl where smsid="&id&"")(0)	
	page=getN("page",1)
	page=int(page)
	if page<=0 then page=1
	pagecount=(count+pagesize-1)\pagesize
	if page>pagecount Then page=pagecount
	rs.move(pagesize*(page-1))
	For i=1 To PageSize
	If rs.eof Then Exit For	
	w fordate2(rs("pltime"))&"发表<br/>　　"   
	w noubb(rs("pl"))&"<br/>-----------<br/>"
	rs.moveNext
	Next
	if page>1 then w "<a href="""&gopage&"page=1"">首页</a>&nbsp;"
	if page-pagecount<0 then w "<a href="""&gopage&"page="&page+1&""">下页</a>&nbsp;"
	if page>1 then w "<a href="""&gopage&"page="&page-1&""">上页</a>&nbsp;"
	if page-pagecount<0 then w "<a href="""&gopage&"page="&pagecount&""">末页</a>"
	if pagecount>1 then w "<br/>第"&page&"页 共"&pagecount&"页<br/>第<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/>页 <a href="""&gopage&"page=$(page)"">跳转</a><br/>"
	Else
	w "暂时没有评论！<br/> "
	end if
	rs.close
	set rs=nothing
	w "<a href='?aid=art&amp;id="&id&"&amp;p="&p&">'>返回原文</a> <a href='?aid=list&amp;p="&p&">'>返回上级栏目</a><br/>"
End Sub
' Wap2.0评论页
Sub showsDiscuss()
	Dim p,rs
	p=getN("p",1)
	w "<title>网友跟贴</title></head><body><div class=""main""><div class=""nav"">-<a href=""?aid=index"">首页</a>-<a href=""?aid=art&amp;id="&id&"&amp;p="&p&""">查看原文</a>-跟贴</div>"&_
	  "发表评论：<br/><form name="""&Time_r&""" action=""?"" method=""get""><input type=""text"" name=""pl"" value="""" maxlength=""200""/>"&_
	  "<input type=""hidden"" name=""aid"" value=""diss""/><input type=""hidden"" name=""id"" value="""&id&"""/><input type=""hidden"" name=""p"" value="""&p&"""/>"&_
	  "<br/><input type=""submit"" value=""提交""/></form><div class=""nav"">【网友评论区】</div>"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open"select id,pl,pltime,ag,da from 74hu_pl where smsid="&id&" order by id desc",conn,3,1
	If Not rs.eof Then
	Dim PageSize,gopage,Count,page,i,pagecount
	PageSize=8
	gopage="?aid=dis&amp;id="&id&"&amp;p="&p&"&amp;"
	Count=conn.execute("Select count(ID) from 74hu_pl where smsid="&id&"")(0)
	page=getN("page",1)
	page=int(page)
	if page<=0 then page=1
	pagecount=(count+pagesize-1)\pagesize
	if page>pagecount Then page=pagecount
	rs.move(pagesize*(page-1))
	For i=1 To PageSize
	If rs.eof Then Exit For	
	w fordate2(rs("pltime"))&"发表<br/>　　"   
	w noubb(rs("pl"))&"<br/>-----------<br/>"
	rs.moveNext
	Next
	if page>1 then w "<a href="""&gopage&"page=1"">首页</a>&nbsp;"
	if page-pagecount<0 then w "<a href="""&gopage&"page="&page+1&""">下页</a>&nbsp;"
	if page>1 then w "<a href="""&gopage&"page="&page-1&""">上页</a>&nbsp;"
	if page-pagecount<0 then w "<a href="""&gopage&"page="&pagecount&""">末页</a>"
	if pagecount>1 then w "<br/>第"&page&"页 共"&pagecount&"页<br/>"&turnpage2("dis","<input name=""p"" type=""hidden"" value="""&p&"""/><input name=""id"" type=""hidden"" value="""&id&"""/>")&"<br/>"
	Else
	w "暂时没有评论！<br/> "
	end if
	rs.close
	set rs=nothing
	w "<a href='?aid=art&amp;id="&id&"&amp;p="&p&">'>返回原文</a> <a href='?aid=list&amp;p="&p&">'>返回上级栏目</a><br/>"
End Sub
' Wap1.0 讨论页
Sub showComment()
	Dim o,rs
	w "<card title='发表评论'>"
	o=getN("o",0)
	if o=0 then
	Dim pl,p
	pl=getD("pl","")
	p=getN("p",1)
	if pl="" then wn "<p>评论内容不能为空！</p></card></wml>"
	if len(pl)>100 Then wn"<p>评论内容最多100字！</p></card></wml>"
	Dim Counts
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open"select * from 74hu_pl",conn,1,3
	rs.addnew
	rs("pl")=pl
	rs("ip")=User_Ip
	rs("smsid")=id
	rs.update
	'更新评论
	Counts=conn.execute("Select count(ID) from 74hu_pl where smsid="&id&"")(0)
	conn.execute("update 74hu_article Set smspin = "&Counts&" where ID="&id)
	else
	Dim ids
	ids=getN("ids",1)
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open"select ag,da from 74hu_pl where id="&ids,conn,1,3
	if o=1 then
	rs("ag")=rs("ag")+1
	else
	rs("da")=rs("da")+1
	end if
	rs.update
	end if
	w "<onevent type='onenterforward'><go href='?aid=dis&amp;id="&id&"&amp;p="&p&"'/></onevent><p>评论发表成功！<br/>"
	rs.close
	Set rs=nothing
	w "<a href='?aid=dis&amp;id="&id&"&amp;p="&p&"'>查看评论</a><br/><a href='?aid=art&amp;id="&id&"&amp;p="&p&">'>返回原文</a><br/>"
End Sub
' Wap2.0讨论页
Sub showsComment()
	w "<title>发表评论</title></head><body><div class=""main"">"
	Dim o,rs
	o=getN("o",0)
	if o=0 then
	Dim pl,p,Counts
	pl=getD("pl","")
	p=getN("p",1)
	if pl="" then wn "评论内容不能为空！</div></body></html>"
	if len(pl)>100 then wn "评论内容最多100字！</div></body></html>"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open"select * from 74hu_pl",conn,1,3
	rs.addnew
	rs("pl")=pl
	rs("ip")=User_Ip
	rs("smsid")=id
	rs.update
	'更新评论
	Counts=conn.execute("Select count(ID) from 74hu_pl where smsid="&id&"")(0)
	conn.execute("update 74hu_article Set smspin = "&Counts&" where ID="&id)
	else
	Dim ids
	ids=getN("ids",1)
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open"select ag,da from 74hu_pl where id="&ids,conn,1,3
	if o=1 then
	rs("ag")=rs("ag")+1
	else
	rs("da")=rs("da")+1
	end if
	rs.update
	end if
	w "评论发表成功！<br/>"
	rs.close
	Set rs=nothing
	w "<a href='?aid=dis&amp;id="&id&"&amp;p="&p&"'>查看评论</a><br/><a href='?aid=art&amp;id="&id&"&amp;p="&p&"'>返回原文</a><br/>"
End Sub
' Wap1.0留言页
Sub showGuest()
	Dim p,act,rs
	p=getN("p",1)
	if p<1 then p=1
	act=request.QueryString("act")
	w "<card title=""客服留言"">"
	if act="view" then
	Dim rsn,rspr
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open"select * from 74hu_guest where ID=" & ID,conn,1,1
	set rsn=Server.CreateObject("ADODB.Recordset")
	rsn.open"select * from 74hu_guest where ID<"& ID &" order by id desc",conn,1,1
	set rspr=Server.CreateObject("ADODB.Recordset")
	rspr.open"select * from 74hu_guest where ID>"& ID &" order by id asc",conn,1,1
	w "<p>-<a href='?aid=index'>首页</a>-<a href='?aid=guest'>客服</a>-查看留言<br/><br/>"
	if rs.EOF then
	w "无此留言！<br/>"
	else
	w "作者："&noubb(rs("name"))&"<br/>"&noubb(rs("text"))&"<br/>时间：" & fordate(rs("HU_time")) & "<br/>"
	if rs("retext")<>"" then w "----------<br/>回复："&noubb(rs("retext"))&"<br/>时间："&fordate(rs("retime"))&"<br/>"
	if rsn.recordcount>0 then w "<a href='?aid=guest&amp;act=view&amp;id=" & rsn("ID") & "&amp;p=" & p & "'>下条</a>&nbsp;"
	if rspr.recordcount>0 then w "<a href='?aid=guest&amp;act=view&amp;id=" & rspr("ID") & "&amp;p=" & p & "'>上条</a>"
	if rsn.recordcount>0 or rspr.recordcount>0 then w "<br/>"
	rsn.close
	set rsn=nothing
	rspr.close
	set rspr=nothing
	end if
	elseif act="add" then
	Dim ss
	randomize timer
	ss=Int((9999)*Rnd +1000)
	w "<p>-<a href='?aid=index'>首页</a>-<a href='?aid=guest'>客服</a>-发表留言<br/><br/>昵称：<br/>"&_
		"<input name=""name"" type=""text"" format=""*M"" emptyok=""true"" maxlength=""10""/><br/>"&_
		"标题：<br/><input name=""title"" type=""text"" format=""*M"" emptyok=""true"" maxlength=""20""/><br/>"&_
		"内容：<br/><input name=""text"" type=""text"" format=""*M"" emptyok=""true"" maxlength=""500""/><br/>"&_
		"联系方式(不公开)：<br/><input name=""lianxi"" type=""text"" format=""*M"" emptyok=""true"" maxlength=""50""/><br/>"&_
		"验证码："&ss&"<br/><input name=""num"" type=""text"" format=""*M"" emptyok=""true"" maxlength=""50""/><br/>"&_
		"<anchor>[提交留言]<go href=""?aid=guest&amp;act=save"" method=""get"" accept-charset=""utf-8"">"&_
		"<postfield name=""name"" value=""$(name)""/>"&_
		"<postfield name=""title"" value=""$(title)""/>"&_
		"<postfield name=""text"" value=""$(text)""/>"&_
		"<postfield name=""lianxi"" value=""$(lianxi)""/>"&_
		"<postfield name=""open"" value=""$(open)""/>"&_
		"<postfield name=""num"" value=""$(num)""/>"&_
		"<postfield name=""num1"" value="""&ss&"""/>"&_
		"</go></anchor><br/>"
	elseif act="save" then
	Dim num,num1
	num=request.QueryString("num")
	num1=request.QueryString("num1")
	if num<>num1 then wn "<p>验证码错误,请返回重试！</p></card></wml>"
	Dim name,title,text,lianxi
	name=getD("name","")
	title=getD("title","")
	text=getD("text","")
	lianxi=getD("lianxi","")
	if name="" or title="" or text="" then wn "<p>昵称或标题内容不能为空！</p></card></wml>"
	w "<onevent type='onenterforward'><go href='?aid=guest'/></onevent><p>"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open"select * from 74hu_guest",conn,1,2
	rs.addnew
	rs("name")=name
	rs("title")=title
	rs("text")=text
	rs("HU_time")=now()
	if lianxi<>"" then rs("lianxi")=lianxi
	rs("agent")=User_Ip
	rs.update
	rs.close
	set rs=Nothing
	w "发表成功，正在返回！<br/>"
	else
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open"select * from 74hu_guest order by id desc",conn,1,1
	If Not rs.eof Then
	Dim PageSize,gopage,Count,page,i,pagecount
	PageSize=10
	gopage="?aid=guest&amp;"
	Count=rs.recordcount
	page=getN("page",1)
	page=int(page)
	if page<=0 or page="" then page=1
	pagecount=(count+pagesize-1)\pagesize
	if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
	w "<p>-<a href='?aid=index'>首页</a>-客服首页<br/><br/>共"&count&"条<a href=""?aid=guest&amp;act=add"">留言</a><br/>"
	For i=1 To PageSize
	If rs.eof Then Exit For
	w "<a href='?aid=guest&amp;act=view&amp;id="&rs("ID")&"&amp;p="&p&"'>"&i+(page-1)*PageSize&"."&noubb(rs("title"))&_
		"</a><br/>[网友:"&noubb(rs("name"))
	if rs("retext")<>"" then 
	w "/已回"
	else
	w "/未回"
	end if
	w "]<br/>"
	rs.moveNext
	Next
	if page-pagecount<0 then w "<a href="""&gopage&"page="&page+1&""">下页</a>&nbsp;"
	if page>1 then w "<a href="""&gopage&"page="&page-1&""">上页</a>"
	if pagecount>1 then w "<br/>"&page&"/"&pagecount&"页<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">跳转</a>"
	Else
	w "<p>还没有留言！<br/>"
	end if
	rs.close
	set rs=nothing
	w "<br/><a href=""?aid=guest&amp;act=add"">我要发表留言</a><br/>"
	end if
End Sub
' Wap2.0留言页
Sub showsGuest()
	Dim p,act,rs
	p=getN("p",1)
	if p<1 then p=1
	act=request.QueryString("act")
	w "<title>客服留言</title></head><body><div class=""main"">"
	if act="view" then
	Dim rsn,rspr
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open"select * from 74hu_guest where ID=" & ID,conn,1,1
	set rsn=Server.CreateObject("ADODB.Recordset")
	rsn.open"select * from 74hu_guest where ID<"& ID &" order by id desc",conn,1,1
	set rspr=Server.CreateObject("ADODB.Recordset")
	rspr.open"select * from 74hu_guest where ID>"& ID &" order by id asc",conn,1,1
	w "<div class=""nav"">-<a href='?aid=index'>首页</a>-<a href='?aid=guest'>客服</a>-查看留言</div><br/>"
	if rs.EOF then
	w "无此留言！<br/>"
	else
	w "作者："&noubb(rs("name"))&"<br/>"&noubb(rs("text")) & "<br/>时间：" & fordate(rs("HU_time")) & "<br/>"
	if rs("retext")<>"" then w "----------<br/>回复："&noubb(rs("retext"))&"<br/>时间："&fordate(rs("retime"))&"<br/>"
	if rsn.recordcount>0 then w "<a href='?aid=guest&amp;act=view&amp;id=" & rsn("ID") & "&amp;p=" & p & "'>下条</a>&nbsp;"
	if rspr.recordcount>0 then w "<a href='?aid=guest&amp;act=view&amp;id=" & rspr("ID") & "&amp;p=" & p & "'>上条</a>"
	if rsn.recordcount>0 or rspr.recordcount>0 then w "<br/>"
	rsn.close
	set rsn=nothing
	rspr.close
	set rspr=nothing
	end if
	elseif act="add" then
	Dim ss
	randomize timer
	ss=Int((9999)*Rnd +1000)
	w "<div class=""nav"">-<a href='?aid=index'>首页</a>-<a href='?aid=guest'>客服</a>-发表留言</div><br/>昵称：<br/>"&_
	 "<form name=""f"&Time_r&""" action=""?"" method=""get""><input name=""name"" type=""text"" /><br/>标题：<br/>"&_
	 "<input name=""title"" type=""text""/><br/>内容：<br/><input name=""text"" type=""text"" /><br/>联系方式(不公开)：<br/>"&_
	 "<input name=""lianxi"" type=""text""/><br/>验证码："&ss&"<br/><input name=""num"" type=""text""/>"&_
	 "<input name=""aid"" type=""hidden"" value=""guest""/><input name=""act"" type=""hidden"" value=""save""/>"&_
	 "<input name=""num1"" type=""hidden"" value="""&ss&"""/><br/><input type=""submit"" value=""提交留言""/></form>"
	elseif act="save" then
	Dim num,num1
	num=request.QueryString("num")
	num1=request.QueryString("num1")
	if num<>num1 then wn "验证码错误,请返回重试！</div></body></html>"
	Dim name,title,text,lianxi
	name=getD("name","")
	title=getD("title","")
	text=getD("text","")
	lianxi=getD("lianxi","")
	if name="" or title="" or text="" then wn "昵称或标题内容不能为空！</div></body></html>"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open"select * from 74hu_guest",conn,1,2
	rs.addnew
	rs("name")=name
	rs("title")=title
	rs("text")=text
	rs("HU_time")=now()
	if lianxi<>"" then rs("lianxi")=lianxi
	rs("agent")=User_Ip
	rs.update
	rs.close
	set rs=Nothing
	r "?aid=guest"
	else
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open"select * from 74hu_guest order by id desc",conn,1,1
	If Not rs.eof Then
	Dim PageSize,gopage,Count,page,i,pagecount
	PageSize=10
	gopage="?aid=guest&amp;"
	Count=rs.recordcount
	page=getN("page",1)
	page=int(page)
	if page<=0 or page="" then page=1
	pagecount=(count+pagesize-1)\pagesize
	if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
	w "<div class=""nav"">-<a href='?aid=index'>首页</a>-客服首页</div><br/>共"&count&"条<a href=""?aid=guest&amp;act=add"">留言</a><br/>"
	For i=1 To PageSize
	If rs.eof Then Exit For
	w "<a href='?aid=guest&amp;act=view&amp;id="&rs("ID")&"&amp;p="&p&"'>"&i+(page-1)*PageSize&"."&noubb(rs("title"))&"</a><br/>"&_
		"[网友:"&noubb(rs("name"))
	if rs("retext")<>"" then 
	w "/已回"
	else
	w "/未回"
	end if
	w "]<br/>"
	rs.moveNext
	Next
	if page-pagecount<0 then w "<a href="""&gopage&"page="&page+1&""">下页</a>&nbsp;"
	if page>1 then w "<a href="""&gopage&"page="&page-1&""">上页</a>"
	if pagecount>1 then w "<br/>"&page&"/"&pagecount&"页"&turnpage2("guest","")
	Else
	w "<p>还没有留言！<br/>"
	end if
	rs.close
	set rs=nothing
	w "<br/><a href=""?aid=guest&amp;act=add"">我要发表留言</a><br/>"
	end if
End Sub
' Wap1.0公告页
Sub showReport()
	Dim rs
	IF  Request.QueryString("action")="view" Then
	w "<card id=""index"" title=""查看公告""><p>"
	Set rs=Server.CreateObject("Adodb.Recordset")
	rs.open "select  * from 74hu_gonggao where id="&id&" order by id desc",conn,1,1
	If Not rs.eof Then
	w "标题:"&noubb(rs("name"))&"<br/>["&fordate(rs("HU_time"))&"]<br/>内容:"&ubbcode(rs("title"))&"<br/>"
	Else
	w "没有这个公告！"
	end if
	w "<br/><a href=""?aid=gonggao"">返回公告中心</a><br/>"
	Rs.close
	set Rs=nothing
	else
	w "<card title=""最新公告""><p>"
	Set Rs = Server.CreateObject("Adodb.Recordset")
	Sql = "SELECT * FROM 74hu_gonggao order by id desc"
	Rs.Open Sql,conn,1,1
	If Not rs.eof Then
	Dim PageSize,i,Count,page,pagecount,gopage
	PageSize=10
	gopage="?aid=gonggao&amp;"
	Count=rs.recordcount
	page=getN("page",1)
	if page<1 then page=1
	page=int(page)
	if page<=0 or page="" then page=1
	pagecount=(count+pagesize-1)\pagesize
	if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
	w "共:"&count&"条公告<br/>"
	For i=1 To PageSize
	If rs.eof Then Exit For
	w "<a href=""?aid=gonggao&amp;action=view&amp;id="&rs("id")&""">"&(i+(page-1)*PageSize)&"."&noubb(Rs("name"))&"</a><br/>"
	rs.moveNext
	Next
	if page-pagecount<0 then w "<a href="""&gopage&"page="&page+1&""">下页</a>"
	if page>1 then w "<a href="""&gopage&"page="&page-1&""">上页</a>"
	if pagecount>1 then w "<br/><b>"&page&"</b>/"&pagecount&"页<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">[GO]</a><br/>"
	Else
	w "暂时没有公告！<br/>"
	end if
	Rs.close
	set Rs=nothing
	end if
End Sub
' Wap2.0公告页
Sub showsReport()
	Dim rs
	IF  Request.QueryString("action")="view" Then
	w "<title>查看公告</title></head><body><div class=""main""><div class=""nav"">-<a href=""?aid=index"">首页</a>-<a href=""?aid=gonggao"">公告</a>-正文</div>"
	Set rs=Server.CreateObject("Adodb.Recordset")
	rs.open "select  * from 74hu_gonggao where id="&id&" order by id desc",conn,1,1
	If Not rs.eof Then
	w "<div class=""tle"">"&noubb(rs("name"))&"</div>["&fordate(rs("HU_time"))&"]<br/>"&ubbcode(rs("title"))&"<br/>"
	Else
	w "没有这个公告！"
	end if
	w "<br/><a href=""?aid=gonggao"">返回公告中心</a><br/>"
	Rs.close
	set Rs=nothing
	else
	w "<title>最新公告</title></head><body><div class=""main"">"
	Set Rs = Server.CreateObject("Adodb.Recordset")
	Sql = "SELECT * FROM 74hu_gonggao order by id desc"
	Rs.Open Sql,conn,1,1
	If Not rs.eof Then
	Dim PageSize,gopage,Count,page,i,pagecount
	PageSize=10
	gopage="?aid=gonggao&amp;"
	Count=rs.recordcount
	page=getN("page",1)
	if page<1 then page=1
	page=int(page)
	if page<=0 or page="" then page=1
	pagecount=(count+pagesize-1)\pagesize
	if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
	w "<div class=""nav"">-<a href=""?aid=index"">首页</a>-公告中心</div>共:"&count&"条公告<br/>"
	For i=1 To PageSize
	If rs.eof Then Exit For
	w "<a href=""?aid=gonggao&amp;action=view&amp;id="&rs("id")&""">"&(i+(page-1)*PageSize)&"."&noubb(Rs("name"))&"</a><br/>"
	rs.moveNext
	Next
	if page-pagecount<0 then w "<a href="""&gopage&"page="&page+1&""">下页</a>"
	if page>1 then w "<a href="""&gopage&"page="&page-1&""">上页</a>"
	if pagecount>1 then w "<br/><b>"&page&"</b>/"&pagecount&"页"&turnpage2("gonggao","")&"<br/>"
	Else
	w "暂时没有公告！<br/>"
	end if
	Rs.close
	set Rs=nothing
	end if
End Sub
' wap1.0列表页
Sub showList()
	Dim act,rs
	act=request.QueryString("act")
	if act<>"" then
	w "<card title='站内排行榜'><p>"
	if act="top" then
	w "-<a href='?aid=index'>首页</a>-站内排行-<a href='?aid=list&amp;act=new'>最新</a><br/>-----------<br/>"
	else
	w "-<a href='?aid=index'>首页</a>-站内最新-<a href='?aid=list&amp;act=top'>排行</a><br/>-----------<br/>"
	end if
	else
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="Select class from 74hu_list where classid="&id
	rs.open sql,conn,3,1 
	if rs.eof then
	rs.close
	set rs=Nothing
	r "?aid=index"
	end if
	Dim classname,sql,PageSize,Count,gopage,page,i,pagecount
	classname=rs("class")
	rs.close
	set rs=Nothing
	w "<card title='"&classname&"-"&waptitle&"'><p>-<a href='?aid=index'>首页</a>-"&classname&"-<a href='?aid=list&amp;act=top'>排行</a><br/>-----------<br/>"
	end if
	Set rs = Server.CreateObject("ADODB.Recordset")
	if act<>"" then
	if act="top" then
	sql="Select top 100 id,title from 74hu_article order by hit*1000+id desc"
	else
	sql="Select top 100 id,title from 74hu_article order by id desc"
	end if
	else
	sql="Select id,title from 74hu_article where classid="&id&" order by id desc"
	end if
	rs.open sql,conn,3,1
	If Not rs.eof then
	if adsetkf("ads1")=1 then
	call adstr(1)
	w "<br/>"
	end if
	PageSize=listnums
	if act<>"" then
	if act="top" then
	gopage="?aid=list&amp;act=top&amp;"
	else
	gopage="?aid=list&amp;act=new&amp;"
	end if
	else
	gopage="?aid=list&amp;id="&id&"&amp;"
	end if
	Count=rs.recordcount
	page=getN("page",1)
	if page<=0 or page="" then page=1
	pagecount=(count+pagesize-1)\pagesize
	if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
	For i=1 To PageSize
	If rs.eof Then Exit For
	w "<a href='?aid=art&amp;id="&rs("id")&"'>"&i+(page-1)*PageSize&"."&noubb(rs("title"))&"</a><br/>"
	rs.moveNext
	Next
	if page-pagecount<0 then w "<a href="""&gopage&"page="&page+1&""">下页</a>&nbsp;"
	if page>1 then w "<a href="""&gopage&"page="&page-1&""">上页</a>"
	if pagecount>1 then w "(<b>"&page&"</b>/"&pagecount&")"&"<br/><input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/>页 <a href="""&gopage&"page=$(page)"">翻页</a><br/>"
	w "[相关内容]<br/>"
	call wendtitle(4,id)
	if adsetkf("ads2")=1 then
	call adstr(2)
	w "<br/>"
	end if
	Else
	w "暂时没有文章！<br/>"
	end if
	rs.close
	set rs=nothing
End Sub
' wap2.0列表页
Sub showsList()
	Dim act,rs
	act=request.QueryString("act")
	if act<>"" then
	w "<title>站内排行榜</title></head><body><div class=""main"">"
	if act="top" then
	w "-<a href='?aid=index'>首页</a>-站内排行-<a href='?aid=list&amp;act=new'>最新</a><br/>-----------<br/>"
	else
	w "-<a href='?aid=index'>首页</a>-站内最新-<a href='?aid=list&amp;act=top'>排行</a><br/>-----------<br/>"
	end if
	else
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="Select class from 74hu_list where classid="&id
	rs.open sql,conn,3,1 
	if rs.eof then
	rs.close
	set rs=Nothing
	r "?aid=index"
	end if
	Dim classname,sql
	classname=rs("class")
	rs.close
	set rs=Nothing
	w "<title>"&classname&"-"&waptitle&"</title></head><body>"&_
	  "<div class=""main"">"&_
	  "<div class=""nav"">-<a href='?aid=index'>首页</a>-"&classname&"-<a href='?aid=list&amp;act=top'>排行</a></div>-----------<br/>"
	end if
	Set rs = Server.CreateObject("ADODB.Recordset")
	if act<>"" then
	if act="top" then
	sql="Select top 100 id,title from 74hu_article order by hit*1000+id desc"
	else
	sql="Select top 100 id,title from 74hu_article order by id desc"
	end if
	else
	sql="Select id,title from 74hu_article where classid="&id&" order by id desc"
	end if
	rs.open sql,conn,3,1
	If Not rs.eof then
	if adsetkf("ads1")=1 then
	call adstr(1)
	w "<br/>"
	end if
	Dim PageSize,gopage,Count,page,i,pagecount
	PageSize=listnums
	if act<>"" then
	if act="top" then
	gopage="?aid=list&amp;act=top&amp;"
	else
	gopage="?aid=list&amp;act=new&amp;"
	end if
	else
	gopage="?aid=list&amp;id="&id&"&amp;"
	end if
	Count=rs.recordcount
	page=getN("page",1)
	if page<=0 or page="" then page=1
	pagecount=(count+pagesize-1)\pagesize
	if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
	For i=1 To PageSize
	If rs.eof Then Exit For
	w "<a href='?aid=art&amp;id="&rs("id")&"'>"&i+(page-1)*PageSize&"."&noubb(rs("title"))&"</a><br/>"
	rs.moveNext
	Next
	if page-pagecount<0 then w "<a href="""&gopage&"page="&page+1&""">下页</a>&nbsp;"
	if page>1 then w "<a href="""&gopage&"page="&page-1&""">上页</a>"
	if pagecount>1 then w "(<b>"&page&"</b>/"&pagecount&")"&"<br/>"&turnpage2("list","<input name=""id"" type=""hidden"" value="""&id&"""/>")&"<br/>"
	w "<div class=""nav"">[相关内容]</div>"
	call wendtitle(4,id)
	if adsetkf("ads2")=1 then
	call adstr(2)
	w "<br/>"
	end if
	Else
	w "暂时没有文章！<br/>"
	end if
	rs.close
	set rs=nothing
End Sub
' Wap1.0地图页
Sub showMap()
	Dim rs,sql
	w "<card title="""&waptitle&"网站地图""><p>"
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="Select * from 74hu_list"
	rs.open sql,conn,1,1 
	if not rs.eof then
	Dim i
	For i=1 to rs.recordcount
	w "<a href=""?aid=list&amp;id="&rs("classid")&""">"&i&"."&rs("class")&"</a><br/>"
	rs.moveNext
	Next
	rs.close
	set rs=nothing
	else
	w "暂时没有<br/> "
	end if
End Sub
' Wap2.0地图页
Sub showsMap()
	w "<title>"&waptitle&"网站地图</title></head><body><div class=""main""><div class=""nav"">-<a href=""?aid=index"">首页</a>-网站地图</div>"
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="Select * from 74hu_list"
	rs.open sql,conn,1,1 
	if not rs.eof then
	Dim i
	For i=1 to rs.recordcount
	w "<a href=""?aid=list&amp;id="&rs("classid")&""">"&i&"."&rs("class")&"</a><br/>"
	rs.moveNext
	Next
	rs.close
	set rs=nothing
	else
	w "暂时没有<br/> "
	end if
End Sub
' Wap1.0书签页
Sub showBookmark()
w "<card title=""保存书签""><p>您可以按以下步骤收藏本站。<br/><br/>1.诺基亚:依次选""操作""-""增加书签"" <br/>"&_
	"2.摩托罗拉:依次选""菜单键""-""书签""-""标记站点""-""保存"" <br/>3.索爱:依次选""更多""-""书签""-""添加书签""-""确定""<br/>"&_
	"4.三星:依次选""上网键""-""收藏夹""-选择一个空的收藏夹地址-确认url地址-输入"""&waptitle&""" <br/>"&_
	"5.松下:选择页面左上角的""菜单""-""书签""-""标记站点""-""保存"" <br/>6.西门子:依次选""上网键""-""收藏夹""-""储存"" <br/>"&_
	"7.NEC:依次选""菜单""-""书签""-""标记站点""-""保存"" <br/>8.LG:依次选""菜单""-""书签""-""标记站点""-""保存"" <br/>"&_
	"9.三菱:按左功能键-""书签""-""添加新书签""-""保存"" <br/>10.海尔:网页浏览状态下长按""*""键-""书签""-""新建""-""编辑""-输入"""&_
	waptitle&"""-""保存"" <br/>11.夏新:访问网站时选中页面左上角-""书签""-""保存"" <br/>12.联想:依次选""网络""-""输入网址""-输入"""&_
	waptitle&"""-""保存"" <br/>13.东信:依次选""选项""-""保存书签"" <br/>14.CECT:依次选""菜单""-""保存书签""-""保存""<br/>"
End Sub
' Wap2.0书签页
Sub showsBookmark()
w "<title>保存书签</title></head><body><div class=""main""><div class=""nav""><a href=""?aid=index"">首页</a>-收藏本站</div>"&_
	"您可以按以下步骤收藏本站。<br/><br/>1.诺基亚:依次选""操作""-""增加书签"" <br/>"&_
	"2.摩托罗拉:依次选""菜单键""-""书签""-""标记站点""-""保存"" <br/>3.索爱:依次选""更多""-""书签""-""添加书签""-""确定""<br/>"&_
	"4.三星:依次选""上网键""-""收藏夹""-选择一个空的收藏夹地址-确认url地址-输入"""&waptitle&""" <br/>"&_
	"5.松下:选择页面左上角的""菜单""-""书签""-""标记站点""-""保存"" <br/>6.西门子:依次选""上网键""-""收藏夹""-""储存"" <br/>"&_
	"7.NEC:依次选""菜单""-""书签""-""标记站点""-""保存"" <br/>8.LG:依次选""菜单""-""书签""-""标记站点""-""保存"" <br/>"&_
	"9.三菱:按左功能键-""书签""-""添加新书签""-""保存"" <br/>10.海尔:网页浏览状态下长按""*""键-""书签""-""新建""-""编辑""-输入"""&_
	waptitle&"""-""保存"" <br/>11.夏新:访问网站时选中页面左上角-""书签""-""保存"" <br/>12.联想:依次选""网络""-""输入网址""-输入"""&_
	waptitle&"""-""保存"" <br/>13.东信:依次选""选项""-""保存书签"" <br/>14.CECT:依次选""菜单""-""保存书签""-""保存""<br/>"
End Sub
' Wap1.0 链接页
Sub showUrl()
	Dim rs
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open"select * from 74hu_gogo Where id="&id,conn,1,1
	if not (rs.bof and rs.eof) then
	Dim tid,url
	tid=rs("id")
	url=noubburl(rs("url"))
	conn.Execute("update 74hu_gogo set tid=tid+1 Where id=" & tid)
	else
	url="?aid=index"
	end if
	rs.close
	set rs=nothing
	w "<card title='正在进入...'><onevent type='onenterforward'><go href='"&noubburl(url)&"'/></onevent>"&_
		"<p align=""left"" mode=""wrap"">如果网页没有自动跳转，请点击<a href="""&noubb(url)&">"">快速进入</a><br/>"
End Sub
' Wap2.0 链接页
Sub showsUrl()
	Dim rs
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open"select * from 74hu_gogo Where id="&id,conn,1,1
	if not (rs.bof and rs.eof) then
	Dim tid,url
	tid=rs("id")
	url=noubburl(rs("url"))
	conn.Execute("update 74hu_gogo set tid=tid+1 Where id=" & tid)
	else
	url="?aid=index"
	end if
	rs.close
	set rs=nothing
	r url
End Sub
' Wap1.0友链页
Sub showLink()
	Dim act,rs,sql,rss,sqll
	act=request.QueryString("act")
	if act="add" then
	w "<card title=""申请友链""><p>"&_
		"网站名称:(3-6字)<br/><input name=""name"&Time_r&""" maxlength=""7"" value=""""/><br/>"&_
		"网站简称:(2汉字)<br/><input name=""namt"&Time_r&""" maxlength=""2"" value=""""/><br/>"&_
		"网址:(需http://)<br/><input name=""url"&Time_r&""" value=""http://""/><br/>"&_
		"网站分类：<select name=""classid"&Time_r&""">"
	Set Rs=server.createobject("adodb.recordset")
	Sql = "select classid,class from 74hu_linkc"
	Rs.open Sql,conn,1,1
	do while not Rs.eof
	w "<option value='"&rs("classid")&"'>"&rs("class")&"</option>"
	rs.movenext
	Loop
	w "</select><br/>"&_
		"网站简介：(50字内)<br/><input name=""jian"&Time_r&""" title=""简介""  value=""暂时没有介绍…"" maxlength=""100""/><br/>"&_
		"<anchor>确定提交<go href=""?aid=link&amp;act=post"" method=""get"" accept-charset=""utf-8"">"&_
		"<postfield name=""name"" value=""$(name"&Time_r&")""/>"&_
		"<postfield name=""namt"" value=""$(namt"&Time_r&")""/>"&_
		"<postfield name=""url"" value=""$(url"&Time_r&")""/>"&_
		"<postfield name=""classid"" value=""$(classid"&Time_r&")""/>"&_
		"<postfield name=""jian"" value=""$(jian"&Time_r&")""/>"&_
		"</go></anchor><br/>"&_
		"<br/>欢迎优秀WAP网站交换链接。"&_
		"<br/>1.合作原则:流量互补,双赢发展,10天没流量首页自动隐藏。"&_
		"<br/>2.流程: "&_
		"<br/>1)提交网站，获取链接地址; "&_
		"<br/>2)将我站的链接放到贵站明显位置。"&_
		"<br/>3)我站人员3个工作日内审核网站，合适网站即可收录。"&_
		"<br/>"&_
		"<br/>申请友情链接前请先在您的网站上做好本站的链接："&_
		"<br/>网站名称："&waptitle&_
		"<br/>做好我站链接后，我们会及时进行审核。<br/>"&_
		"<a href='?aid=link'>返回友链首页</a><br/>"
	rs.close
	set rs=nothing
	elseif act="go" then
	Dim yourip,sss,ips,cache_ip,one_ip,all_s,k_ip,i_ip,del_time,ip_time,temp_s
	On Error Resume Next
	Server.ScriptTimeOut=9999999
	yourip=Request.ServerVariables("HTTP_X_UP_CALLING_LINE_ID")
	if yourip="" then yourip=Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
	if yourip="" then yourip=Request.ServerVariables("REMOTE_ADDR")
	sss=180
	ips=500
	cache_ip=Application("cache_ip")
	if cache_ip="" then cache_ip="|"
	one_ip=split(cache_ip,"|")
	all_s=ubound(one_ip)
	for k_ip=0 to all_s
	if yourip=one_ip(k_ip) then
	i_ip=k_ip:ip_time=one_ip(k_ip+1):Exit for
	else
	i_ip=0:ip_time="2000-10-10 10:10:10"
	end if
	next
	del_time=DATEDIFF("s",ip_time,now())
	if i_ip<all_s and sss>del_time then
	r "?aid=index"
	end if
	if all_s>ips*2 Then
	Application.Lock
	Application("cache_ip")="|"
	Application.UnLock
	else
	if i_ip=0 then
	temp_s=cache_ip&yourip&"|"&now()&"|"
	else
	Dim text_1,num_1,num_2,num_3,num_4,text_2,text_3,text_4
	text_1="|"&yourip&"|"&ip_time&"|"
	num_1=len(cache_ip)
	num_2=len(text_1)
	num_3=instr(cache_ip,text_1)
	num_4=num_1-num_2-num_3+1
	text_2=left(cache_ip,num_3)
	text_3=right(cache_ip,num_4)
	text_4=yourip&"|"&now()&"|"
	temp_s=text_2&text_4&text_3
	end if
	Application.Lock
	Application("cache_ip")=temp_s
	Application.UnLock
	end if
	set Rs=Server.CreateObject("ADODB.Recordset")
	Sql="select ID,HU_in,HU_time from 74hu_link Where ID="&ID
	Rs.open Sql,conn,1,3
	Rs("HU_in")=Rs("HU_in")+1
	Rs("HU_time")=now()
	Rs.update()
	rs.close
	set rs=nothing
	r "?aid=index"
	elseif act="view" then
	set Rs=Server.CreateObject("ADODB.Recordset")
	Sql="select * from 74hu_link Where id="&id
	Rs.open Sql,conn,1,3
	If Not rs.eof	Then
	Rs("HU_out")=Rs("HU_out")+1
	Rs("OUTtime")=now()
	Rs.update()
	Else
	w "?aid=index"
	End If
	w "<card title='"&noubb(rs("name"))&"' ontimer='"&noubburl(rs("url"))&"'><timer value='1'/><p>"&_
		"正在跳转到“"&noubb(rs("name"))&"”,<br/>请稍候...<a href='"&noubburl(rs("url"))&"'>快速进入</a><br/>"&_
		"网站介绍："&noubb(rs("jian"))&"<br/><br/>"
	rs.close
	set rs=nothing
	elseif act="post" then
	Dim classid,name,namt,url,jian
	classid=Request.QueryString("classid")
	name=getD("name","")
	namt=getD("namt","")
	url=LCase(getD("url",""))
	jian=getD("jian","")
	if session("name")=1 then
	wn "<card title=""重复申请""><p>你刚才已申请过了！请不要重复申请！<br/>"
	else
	if name="" or namt="" or url="" or jian="" or classid="" or isnumeric(classid)=false then
	wn "<card title=""出错了吧""><p>各项都要填写,不能为空！<br/>"
	else
	Set RSS=server.createobject("adodb.recordset")
	Sqll="select * from 74hu_ad"
	RSS.open sqll,conn,1,1
	active=rss("active")
	rss.close
	set rss=nothing
	Set RS=server.createobject("adodb.recordset")
	Sql="select * from 74hu_link"
	RS.open sql,conn,1,3
	RS.addnew
	RS("name")=name
	RS("namt")=namt
	RS("url")=url
	RS("classid")=classid
	RS("jian")=jian
	RS("active")=active
	RS.update
	session.timeout=1
	session("name")=1
	end if
	end if
	w "<card title='申请友链成功' ontimer='?aid=link&amp;act=you'><timer value='1'/><p>申请友链成功<br/>"
	rs.close
	set rs=nothing
	elseif act="list" then
	Dim add
	add=request.QueryString("class")
	if add="" or IsNumeric(add)=false then
	r "?aid=index"
	else
	Set rss=Server.CreateObject("ADODB.Recordset")
	sqll="Select * from 74hu_linkc where classid="&add
	rss.open sqll,conn,1,1 
	if not rss.eof then
	Dim classname
	classname=rss("class")
	end if
	rss.close
	set rss=nothing
	w "<card title="""&classname&"网站""><p>=" &classname&"网站=<br/>"
	Dim PageSize,gopage,Count,page,i,pagecount
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="Select * from 74hu_link where classid="&add&" And Active=0 and del=0  order by HU_time desc"
	rs.open sql,conn,1,1 
	If Not rs.eof Then
	PageSize=15
	gopage="?aid=link&amp;act=list&amp;class="&ADD&"&amp;"
	Count=conn.execute("Select count(ID) from 74hu_link where classid="&add&"")(0)
	page=getN("page",1)
	if page<=0 then page=1
	pagecount=(count+pagesize-1)\pagesize
	if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
	For i=1 To PageSize
	If rs.eof Then Exit For
	w ""&i+(page-1)*PageSize&".<a href=""?aid=link&amp;act=view&amp;class="&rs("classid")&"&amp;id="&rs("id")&""">"&noubb(rs("name"))&"</a><br/>"
	rs.moveNext
	Next
	if page-pagecount<0 then w "<a href="""&gopage&"page="&page+1&""">下页</a>"
	if page>1 then w "<a href="""&gopage&"page="&page-1&""">上页</a>"
	if pagecount>1 then w "<br/><b>"&page&"</b>/"&pagecount&"页<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">>></a><br/>"
	Else
	w "暂时没有添加！<br/>"
	end if
	w "<br/><a href='?aid=link'>返回友链首页</a><br/>"
	rs.close
	set rs=nothing
	end if
	elseif act="you" then
	w "<card title=""申请友链成功""><p>"
	set Rss=Server.CreateObject("ADODB.Recordset")
	Sqll="select Active from 74hu_ad"
	Rss.open Sqll,conn,1,1
	Dim Active
	Active=Rss("Active")
	rss.close
	set rss=nothing
	set Rs=Server.CreateObject("ADODB.Recordset")
	Sql="select top 1 * from 74hu_link order by id desc"
	RS.open Sql,conn,1,1
	w "添加友链地址成功，"
	if Active=1 then
	w "请等待站长审核，审核通过后才会显示<br/>"
	else
	w "你的友链已经显示出来!<br/>"
	End if
	w "贵站返回我站的链接地址是:http://"&wapurl&"/?aid=link&amp;act=go&amp;id="&Rs("ID")&"<br/>网站名称:"&waptitle&"<br/>"
	rs.close
	set rs=nothing
	else
	w "<card title=""友情链接""><p>----动态友链----<br/>"
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select linkactive from 74hu_ad"
	rs.open sql,conn,1,1
	if not rs.eof then
	DIm getday
	getday=rs("linkactive")
	end if
	rs.close
	set rs=nothing
	set rs=Server.CreateObject("ADODB.Recordset")
	Sql="select ID,name,classid from 74hu_link Where active=0 and del=0 and datediff('d', HU_time, now())<"&getday&" order by HU_time desc"
	rs.open Sql,conn,1,1
	If Not rs.eof	Then
	PageSize=30
	gopage="?aid=link&amp;"
	Count=rs.recordcount
	page=getN("page",1)
	if page<=0 then page=1
	pagecount=(count+pagesize-1)\pagesize
	if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
	For i=1 To PageSize
	If rs.eof Then Exit For
	w "<a href=""?aid=link&amp;act=view&amp;id="&rs("id")&"&amp;class="&rs("classid")&""">"&noubb(rs("name"))&"</a><br/>" 
	rs.moveNext
	Next
	if page-pagecount<0 then w "<a href="""&gopage&"page="&page+1&""">下页</a>"
	if page>1 then w "<a href="""&gopage&"page="&page-1&""">上页</a>"
	if pagecount>1 then w "<br/><b>"&page&"</b>/"&pagecount&"页<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">跳转</a>"
	Else
	w "暂时没有友链！<br/>"
	end if
	rs.close
	set rs=nothing
	w "----网站分类----<br/>"
	set rs = server.createobject("adodb.recordset")
	sql="select linkindex from 74hu_ad"
	rs.open sql,conn,1,1
	if rs.eof then
	rs.close
	set rs=nothing
	wn "<p>资料没有配置！</p></card></wml>"
	end if
	linkindex=rs("linkindex")
	rs.close
	set rs=nothing
	if linkindex=0 then
	Dim rs1,sql1
	set rs1 = server.createobject("adodb.recordset")
	sql1="select * from 74hu_linkc order by pid asc"
	rs1.open sql1,conn,1,1
	if rs1.eof then
	w "暂无分类<br/>"
	else
	i=1
	do while not rs1.eof
	Dim sqls
	set Rss=Server.CreateObject("ADODB.Recordset")
	Sqls="select top 4 ID,namt,classid from 74hu_link Where Active=0 and del=0 and classid="&rs1("classid")&" order by HU_time desc"
	rss.open Sqls,conn,1,1
	w "【<a href=""?aid=link&amp;act=list&amp;class="&rs1("classid")&""">"&noubb(rs1("Class"))&"</a>】"
	If rss.eof then
	w "暂时还没有<br/>"& chr(13)
	Else
	for i=1 to 6
	w "<a href='?aid=link&amp;act=view&amp;class="&rss("classid")&"&amp;id="&rss("id")&"'>"&noubb(rss("namt"))&"</a> "
	rss.Movenext
	if rss.EOF then Exit for
	Next
	w "<br/>"
	End if
	rss.close
	set rss=nothing
	i=i+1
	rs1.movenext
	loop
	end if
	rs1.close
	set rs1=nothing
	else
	set rs = server.createobject("adodb.recordset")
	sql="select * from 74hu_linkc order by pid asc"
	rs.open sql,conn,1,1
	if not (rs.bof and rs.eof)  then
	For i=1 to rs.RecordCount
	If Rs.Eof Then
	exit For
	End If
	Dim br
	if rs("br")="1" then
	br="<br/>"
	else
	br=""
	end if
	w "<a href=""?aid=link&amp;act=list&amp;class="&rs("classid")&""">"&noubb(rs("class"))&"</a> "& br &"" 
	Rs.MoveNext
	Next
	end if
	rs.close
	set rs=nothing
	end if
	w "<a href='?aid=link&amp;act=add'>&gt;&gt;友链合作申请</a>"
	end if
End Sub
' Wap2.0友链页
Sub showsLink()
	Dim act,rs,sql,rss
	act=request.QueryString("act")
	if act="add" then
	w "<title>申请友链</title></head><body><div class=""main""><form name=""f"&Time_r&""" action=""?"" method=""get"">"&_
		"网站名称:(3-6字)<br/><input name=""name"" maxlength=""7"" value=""""/><br/>"&_
		"网站简称:(2汉字)<br/><input name=""namt"" maxlength=""2"" value=""""/><br/>"&_
		"网址:(需http://)<br/><input name=""url"" value=""http://""/><br/>"&_
		"网站分类：<select name=""classid"">"
	Set Rs=server.createobject("adodb.recordset")
	Sql = "select classid,class from 74hu_linkc"
	Rs.open Sql,conn,1,1
	do while not Rs.eof
	w "<option value="""&rs("classid")&""">"&rs("class")&"</option>"
	rs.movenext
	Loop
	w "</select><br/>"&_
		"网站简介：(50字内)<br/><input name=""jian"" value=""暂时没有介绍…"" maxlength=""100"" accept-charset=""utf-8""/><br/>"&_
		"<input type=""hidden"" name=""aid"" value=""link""/><input type=""hidden"" name=""act"" value=""post""/>"&_
		"<input type=""submit"" value=""确定提交""/></form>"&_
		"<br/>"&_
		"<br/>欢迎优秀WAP网站交换链接。"&_
		"<br/>1.合作原则:流量互补,双赢发展,10天没流量首页自动隐藏。"&_
		"<br/>2.流程: "&_
		"<br/>1)提交网站，获取链接地址; "&_
		"<br/>2)将我站的链接放到贵站明显位置。"&_
		"<br/>3)我站人员3个工作日内审核网站，合适网站即可收录。"&_
		"<br/>"&_
		"<br/>申请友情链接前请先在您的网站上做好本站的链接："&_
		"<br/>网站名称："&waptitle&_
		"<br/>做好我站链接后，我们会及时进行审核。<br/>"&_
		"<a href='?aid=link'>返回友链首页</a><br/>"
	rs.close
	set rs=nothing
	elseif act="go" then
	On Error Resume Next
	Dim yourip,sss,ips,cache_ip,one_ip,all_s,k_ip,i_ip,del_time,ip_time,temp_s
	Server.ScriptTimeOut=9999999
	yourip=Request.ServerVariables("HTTP_X_UP_CALLING_LINE_ID")
	if yourip="" then yourip=Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
	if yourip="" then yourip=Request.ServerVariables("REMOTE_ADDR")
	sss=180
	ips=500
	cache_ip=Application("cache_ip")
	if cache_ip="" then cache_ip="|"
	one_ip=split(cache_ip,"|")
	all_s=ubound(one_ip)
	for k_ip=0 to all_s
	if yourip=one_ip(k_ip) then
	i_ip=k_ip:ip_time=one_ip(k_ip+1):Exit for
	else
	i_ip=0:ip_time="2000-10-10 10:10:10"
	end if
	next
	del_time=DATEDIFF("s",ip_time,now())
	if i_ip<all_s and sss>del_time then
	r "?aid=index"
	end if
	if all_s>ips*2 Then
	Application.Lock
	Application("cache_ip")="|"
	Application.UnLock
	else
	if i_ip=0 then
	temp_s=cache_ip&yourip&"|"&now()&"|"
	else
	Dim text_1,num_1,num_2,num_3,num_4,text_2,text_3,text_4
	text_1="|"&yourip&"|"&ip_time&"|"
	num_1=len(cache_ip)
	num_2=len(text_1)
	num_3=instr(cache_ip,text_1)
	num_4=num_1-num_2-num_3+1
	text_2=left(cache_ip,num_3)
	text_3=right(cache_ip,num_4)
	text_4=yourip&"|"&now()&"|"
	temp_s=text_2&text_4&text_3
	end if
	Application.Lock
	Application("cache_ip")=temp_s
	Application.UnLock
	end if
	set Rs=Server.CreateObject("ADODB.Recordset")
	Sql="select ID,HU_in,HU_time from 74hu_link Where ID="&ID
	Rs.open Sql,conn,1,3
	Rs("HU_in")=Rs("HU_in")+1
	Rs("HU_time")=now()
	Rs.update()
	rs.close
	set rs=nothing
	r "?aid=index"
	elseif act="view" then
	set Rs=Server.CreateObject("ADODB.Recordset")
	Sql="select * from 74hu_link Where id="&id
	Rs.open Sql,conn,1,3
	If Not rs.eof	Then
	Rs("HU_out")=Rs("HU_out")+1
	Rs("OUTtime")=now()
	Rs.update()
	Else
	r "?aid=index"
	End If
	w "<title>"&noubb(rs("name"))&"</title></head><body><div class=""main"">"&_
		"<div class=""nav"">-<a href=""?aid=index"">首页</a>-<a href=""?aid=link"">友链</a>-查看网站</div>"&_
		"站名："&noubb(rs("name"))&"<br/>"&_
		"介绍："&noubb(rs("jian"))&"<br/>"&_
		"- > <a href="""&noubburl(rs("url"))&""">访问网站</a><br/><br/>"
	rs.close
	set rs=nothing
	elseif act="post" then
	Dim classid,name,namt,url,jian
	classid=Request.QueryString("classid")
	name=getD("name","")
	namt=getD("namt","")
	url=LCase(getD("url",""))
	jian=getD("jian","")
	if session("name")=1 then
	wn "<title>重复申请</title></head><body><div class=""main"">你刚才已申请过了！请不要重复申请！,<a href=""?aid=link&amp;act=post"">返回</a></div></body></html>"
	else
	if name="" or namt="" or url="" or jian="" or classid="" or isnumeric(classid)=false then
	wn "<title>出错了吧</title></head><body><div class=""main"">各项都要填写,不能为空！,<a href=""?aid=link&amp;act=post"">返回</a></div></body></html>"
	else
	Set RSS=server.createobject("adodb.recordset")
	Sqll="select * from 74hu_ad"
	RSS.open sqll,conn,1,1
	active=rss("active")
	rss.close
	set rss=nothing
	Set RS=server.createobject("adodb.recordset")
	Sql="select * from 74hu_link"
	RS.open sql,conn,1,3
	RS.addnew
	RS("name")=name
	RS("namt")=namt
	RS("url")=url
	RS("classid")=classid
	RS("jian")=jian
	RS("active")=active
	RS.update
	session.timeout=1
	session("name")=1
	end if
	end if
	w "<title>申请友链成功</title></head><body><div class=""main"">申请友链成功,<a href=""?aid=link&amp;act=you"">查看回链</a><br/>"
	rs.close
	set rs=nothing
	elseif act="list" then
	add=request.QueryString("class")
	if add="" or IsNumeric(add)=false then
	r "?aid=index"
	else
	Dim sqll
	Set rss=Server.CreateObject("ADODB.Recordset")
	sqll="Select * from 74hu_linkc where classid="&add
	rss.open sqll,conn,1,1 
	if not rss.eof then
	Dim classname
	classname=rss("class")
	end if
	rss.close
	set rss=nothing
	Dim PageSize,gopage,Count,page,i,pagecount
	w "<title>"&classname&"网站</title></head><body><div class=""main""><div class=""nav"">-<a href=""?aid=index"">首页</a>-<a href=""?aid=link"">友链</a>-"&classname&"网站</div>"
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="Select * from 74hu_link where classid="&add&" And Active=0 and del=0  order by HU_time desc"
	rs.open sql,conn,1,1 
	If Not rs.eof Then
	PageSize=15
	gopage="?aid=link&amp;act=list&amp;class="&ADD&"&amp;"
	Count=conn.execute("Select count(ID) from 74hu_link where classid="&add&"")(0)
	page=getN("page",1)
	if page<=0 then page=1
	pagecount=(count+pagesize-1)\pagesize
	if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
	For i=1 To PageSize
	If rs.eof Then Exit For
	w ""&i+(page-1)*PageSize&".<a href=""?aid=link&amp;act=view&amp;class="&rs("classid")&"&amp;id="&rs("id")&""">"&noubb(rs("name"))&"</a><br/>"
	rs.moveNext
	Next
	if page-pagecount<0 then w "<a href="""&gopage&"page="&page+1&""">下页</a>"
	if page>1 then w "<a href="""&gopage&"page="&page-1&""">上页</a>"
	if pagecount>1 then w "<br/><b>"&page&"</b>/"&pagecount&"页<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">>></a><br/>"
	Else
	w "暂时没有添加！<br/>"
	end if
	w "<br/><a href=""?aid=link"">返回友链首页</a><br/>"
	rs.close
	set rs=nothing
	end if
	elseif act="you" then
	w "<title>申请友链成功</title></head><body><div class=""main"">"
	set Rss=Server.CreateObject("ADODB.Recordset")
	Sqll="select Active from 74hu_ad"
	Rss.open Sqll,conn,1,1
	Dim Active
	Active=Rss("Active")
	rss.close
	set rss=nothing
	set Rs=Server.CreateObject("ADODB.Recordset")
	Sql="select top 1 * from 74hu_link order by id desc"
	RS.open Sql,conn,1,1
	w "添加友链地址成功，"
	if Active=1 then
	w "请等待站长审核，审核通过后才会显示<br/>"
	else
	w "你的友链已经显示出来!<br/>"
	End if
	w "贵站返回我站的链接地址是:http://"&wapurl&"/?aid=link&amp;act=go&amp;id="&Rs("ID")&"<br/>网站名称:"&waptitle&"<br/>"
	rs.close
	set rs=nothing
	else
	w "<title>友情链接</title></head><body><div class=""main""><div class=""nav"">动态友链:</div>"
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select linkactive from 74hu_ad"
	rs.open sql,conn,1,1
	if not rs.eof then
	Dim getday
	getday=rs("linkactive")
	end if
	rs.close
	set rs=nothing
	set rs=Server.CreateObject("ADODB.Recordset")
	Sql="select ID,name,classid from 74hu_link Where active=0 and del=0 and datediff('d', HU_time, now())<"&getday&" order by HU_time desc"
	rs.open Sql,conn,1,1
	If Not rs.eof Then
	PageSize=30
	gopage="?aid=link&amp;"
	Count=rs.recordcount
	page=getN("page",1)
	if page<=0 then page=1
	pagecount=(count+pagesize-1)\pagesize
	if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
	For i=1 To PageSize
	If rs.eof Then Exit For
	w "<a href=""?aid=link&amp;act=view&amp;id="&rs("id")&"&amp;class="&rs("classid")&""">"&noubb(rs("name"))&"</a><br/>" 
	rs.moveNext
	Next
	if page-pagecount<0 then w "<a href="""&gopage&"page="&page+1&""">下页</a>"
	if page>1 then w "<a href="""&gopage&"page="&page-1&""">上页</a>"
	if pagecount>1 then w "<br/><b>"&page&"</b>/"&pagecount&"页<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">跳转</a>"
	Else
	w "暂时没有友链！<br/>"
	end if
	rs.close
	set rs=nothing
	w "<div class=""nav"">网站分类:</div>"
	set rs = server.createobject("adodb.recordset")
	sql="select linkindex from 74hu_ad"
	rs.open sql,conn,1,1
	if rs.eof then
	rs.close
	set rs=nothing
	wn "资料没有配置！</div></body></html>"
	end if
	Dim linkindex
	linkindex=rs("linkindex")
	rs.close
	set rs=nothing
	if linkindex=0 then
	Dim rs1,sql1
	set rs1 = server.createobject("adodb.recordset")
	sql1="select * from 74hu_linkc order by pid asc"
	rs1.open sql1,conn,1,1
	if rs1.eof then
	w "暂无分类<br/>"
	else
	i=1
	do while not rs1.eof
	set Rss=Server.CreateObject("ADODB.Recordset")
	Sqls="select top 4 ID,namt,classid from 74hu_link Where Active=0 and del=0 and classid="&rs1("classid")&" order by HU_time desc"
	rss.open Sqls,conn,1,1
	w "【<a href=""?aid=link&amp;act=list&amp;class="&rs1("classid")&""">"&noubb(rs1("Class"))&"</a>】"
	If rss.eof then
	w "暂时还没有<br/>"
	Else
	for i=1 to 6
	w "<a href='?aid=link&amp;act=view&amp;class="&rss("classid")&"&amp;id="&rss("id")&"'>"&noubb(rss("namt"))&"</a> "
	rss.Movenext
	if rss.EOF then Exit for
	Next
	w "<br/>"
	End if
	rss.close
	set rss=nothing
	i=i+1
	rs1.movenext
	loop
	end if
	rs1.close
	set rs1=nothing
	else
	set rs = server.createobject("adodb.recordset")
	sql="select * from 74hu_linkc order by pid asc"
	rs.open sql,conn,1,1
	if not (rs.bof and rs.eof)  then
	For i=1 to rs.RecordCount
	If Rs.Eof Then
	exit For
	End If
	Dim br
	if rs("br")="1" then
	br="<br/>"
	else
	br=""
	end if
	w "<a href=""?aid=link&amp;act=list&amp;class="&rs("classid")&""">"&noubb(rs("class"))&"</a> "& br &"" 
	Rs.MoveNext
	Next
	end if
	rs.close
	set rs=nothing
	end if
	w "<a href='?aid=link&amp;act=add'>&gt;&gt;友链合作申请</a>"
	end if
End Sub
%>
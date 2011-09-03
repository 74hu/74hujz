<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<%
response.write "<card title=""七色虎建站系统"">"&_
"<p><img src=""logo.png"" alt=""七色虎建站系统""/>"&_
"<br/>欢迎管理员:"&keyuser&"<br/>"&_
"【建站帮助】必看<br/>"&_
"<a href=""faq.asp?sid="&sid&""">用户手册</a> "&_
"<a href=""glyh.asp?sid="&sid&""">官方帮助</a><br/>"&_
"【网站管理】<br/>"&_
"<a href=""class.asp?sid="&sid&""">设计中心</a> "&_
"<a href=""config.asp?sid="&sid&""">网站配置</a><br/>"&_
"【文章管理】<br/>"&_
"<a href=""art/index.asp?sid="&sid&""">文章管理</a> "&_
"<a href=""74hu/text.asp?sid="&sid&""">复制文章</a><br/>"&_
"<a href=""art/wzcl.asp?sid="&sid&""">全部文章</a> "&_
"<a href=""art/wzclass.asp?sid="&sid&""">分类管理</a><br/>"&_
"<a href=""art/tjwz.asp?sid="&sid&""">添加分类</a> "&_
"<a href=""art/adminpl.asp?sid="&sid&""">文章评论</a><br/>"&_
"【站长工具】<br/>"&_
"<a href=""74hu/files.asp?sid="&sid&""">文件管理</a> "&_
"<a href=""74hu/wmltext.asp?sid="&sid&""">自写页面</a><br/>"&_
"<a href=""bak.asp?sid="&sid&""">数据备份</a> "&_
"<a href=""74hu/word.asp?sid="&sid&""">简繁互换</a><br/>"&_
"<a href=""74hu/iplock.asp?sid="&sid&""">管理ＩＰ</a> "&_
"<a href=""74hu/sql.asp?sid="&sid&""">注入攻击</a><br/>"&_
"【高级管理】<br/>"&_
"<a href=""ggao.asp?sid="&sid&""">广告管理</a> "&_
"<a href=""gonggo.asp?sid="&sid&""">公告管理</a><br/>"&_
"<a href=""link/mymin_index.asp?sid="&sid&""">友链管理</a> "&_
"<a href=""lygl/index.asp?sid="&sid&""">留言管理</a><br/>"&_
"<a href=""74hu/index.asp?sid="&sid&""">站长工具</a> "&_
"<a href=""lltj.asp?sid="&sid&""">流量统计</a><br/>"&_
"<a href=""guanli.asp?sid="&sid&""">管理设定</a> "&_
"<a href=""logout.asp?sid="&sid&""">退出管理</a><br/>"&_
"</p></card></wml>"%><%call CloseConn%>
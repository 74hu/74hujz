<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%call head%><card title="常见问题帮助"><p><%

sip=Request.querystring("p")
if isnumeric(sip)=false or isnull(sip) or sip="" then
sip="0"
end if
if sip="0" then
%>
<a href="?p=1&amp;sid=<%=sid%>">1、发表文章</a><br/>
<a href="?p=2&amp;sid=<%=sid%>">2、上传文件</a><br/>
<a href="?p=3&amp;sid=<%=sid%>">3、修改密码</a><br/>
<a href="?p=4&amp;sid=<%=sid%>">4、添加广告</a><br/>
<a href="?p=5&amp;sid=<%=sid%>">8、管理评论</a><br/>
<a href="?p=6&amp;sid=<%=sid%>">6、管理留言</a><br/>
<a href="?p=7&amp;sid=<%=sid%>">7、管理友链</a><br/>
<a href="?p=8&amp;sid=<%=sid%>">8、数据库备份</a><br/>
<a href="?p=9&amp;sid=<%=sid%>">9、网站安全防护</a><br/>
<a href="?p=10&amp;sid=<%=sid%>">10、发布公告</a>~new<br/>
<a href="?p=11&amp;sid=<%=sid%>">11、首页倒计时</a>~new<br/>
<%elseif sip="1" then%>
1、如何发表文章？<br/>
<br/>
分类管理-(添加分类)-点击分类-添加文章<br/>
<br/><anchor>返回问题中心<prev/></anchor>
<%elseif sip="2" then%>
2、如何上传文件？<br/>
<br/>
文件管理-WAP2.0上传-选择文件-点击上传-得到文件地址<br/>
<br/><anchor>返回问题中心<prev/></anchor>
<%elseif sip="3" then%>
3、如何修改密码？<br/>
<br/>
管理设定-点击管理账号-编辑-修改<br/>
<br/><anchor>返回问题中心<prev/></anchor>
<%elseif sip="4" then%>
4、如何添加广告？<br/>
<br/>
广告管理-选择类别-添加广告<br/>
<br/><anchor>返回问题中心<prev/></anchor>
<%elseif sip="5" then%>
5、如何管理评论？<br/>
<br/>
文章评论-删除<br/>
<br/><anchor>返回问题中心<prev/></anchor>
<%elseif sip="6" then%>
6、如何管理留言？<br/>
<br/>
留言管理-点击留言标题-回复留言<br/>
<br/><anchor>返回问题中心<prev/></anchor>
<%elseif sip="7" then%>
7、如何管理友链？<br/>
<br/>
默认已添加网站类别，管理只要进入友链审核<br/>
<br/>
友链管理-友链审核-审核或删除<br/>
<br/><anchor>返回问题中心<prev/></anchor>
<%elseif sip="8" then%>
8、如何备份数据库和恢复数据库？<br/>
<br/>
数据备份-备份或恢复<br/>
<br/><anchor>返回问题中心<prev/></anchor>
<%elseif sip="9" then%>
9、如何进行网站安全管理和防护？<br/>
<br/>
注入攻击-选择SQL注入攻击或后台登陆攻击-记录不法IP地址<br/>
管理ＩＰ-添加IP地址-封锁IP<br/>
注：慎用，部分用户不能访问网站<br/>
<br/><anchor>返回问题中心<prev/></anchor>
<%elseif sip="10" then%>
10、如何发布公告和展示公告？<br/>
<br/>
发布公告：公告管理-发布公告<br/>
修改公告：公告管理-(点击公告)，选择[编辑]或[删除]<br/>
展示公告：网站配置-全站显示公告-显示或不显示<br/>
<br/><anchor>返回问题中心<prev/></anchor>
<%elseif sip="11" then%>
11、如何添加首页倒计时功能？<br/>
<br/>
网站配置-(首页倒计时)-填写时间<br/>
注：格式为2008-8-8<br/>
要添加相应的倒计时名称：<br/>
网站配置-(倒计时名称)-填写名称<br/>
注：显示为“距XX还有N天”<br/>
<br/>
那如何去掉首页倒计时？<br/>
在(首页倒计时)处保持空白即可<br/>
<br/><anchor>返回问题中心<prev/></anchor>
<%end if%>
<br/><a href="index.asp?sid=<%=sid%>">后台管理</a>
</p></card></wml>
<%call CloseConn%>
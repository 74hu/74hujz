<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%call head%><card title="常见问题帮助"><p><%
Dim sip
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
<a href="?p=10&amp;sid=<%=sid%>">10、发布公告</a><br/>
<a href="?p=11&amp;sid=<%=sid%>">11、首页倒计时</a><br/>
<a href="?p=12&amp;sid=<%=sid%>">12、UBB语言</a>~<br/>
<a href="?p=13&amp;sid=<%=sid%>">13、第三方流量统计</a>~<br/>
<a href="?p=14&amp;sid=<%=sid%>">14、WML标签</a>~<br/>
<a href="?p=15&amp;sid=<%=sid%>">15、网站底部编辑</a>~<br/>
<a href="?p=16&amp;sid=<%=sid%>">16、切换网站风格</a>~<br/>
<a href="?p=17&amp;sid=<%=sid%>">17、过滤敏感词</a>~<br/>
<a href="?p=18&amp;sid=<%=sid%>">18、首页标题字数</a>~<br/>
<br/>
温馨提示：<br/>
还有问题或建议可以到官网反馈<br/>
<%elseif sip="1" then%>
1、如何发表文章？<br/>
<br/>
分类管理-(添加分类)-点击分类-添加文章<br/>
<%elseif sip="2" then%>
2、如何上传文件？<br/>
<br/>
文件管理-WAP2.0上传-选择文件-点击上传-得到文件地址<br/>
<%elseif sip="3" then%>
3、如何修改密码？<br/>
<br/>
管理设定-点击管理账号-编辑-修改<br/>
<%elseif sip="4" then%>
4、如何添加广告？<br/>
<br/>
广告管理-选择类别-添加广告<br/>
<%elseif sip="5" then%>
5、如何管理评论？<br/>
<br/>
文章评论-删除<br/>
<%elseif sip="6" then%>
6、如何管理留言？<br/>
<br/>
留言管理-点击留言标题-回复留言<br/>
<%elseif sip="7" then%>
7、如何管理友链？<br/>
<br/>
默认已添加网站类别，管理只要进入友链审核<br/>
<br/>
友链管理-友链审核-审核或删除<br/>
<%elseif sip="8" then%>
8、如何备份数据库和恢复数据库？<br/>
<br/>
数据备份-备份或恢复<br/>
<%elseif sip="9" then%>
9、如何进行网站安全管理和防护？<br/>
<br/>
注入攻击-选择SQL注入攻击或后台登陆攻击-记录不法IP地址<br/>
管理ＩＰ-添加IP地址-封锁IP<br/>
注：慎用，部分用户不能访问网站<br/>
<%elseif sip="10" then%>
10、如何发布公告和展示公告？<br/>
<br/>
发布公告：公告管理-发布公告<br/>
修改公告：公告管理-(点击公告)，选择[编辑]或[删除]<br/>
展示公告：网站配置-全站显示公告-显示或不显示<br/>
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
<%elseif sip="12" then%>
12、本系统支持那些ubb语言？<br/>
<br/>----分页----<br/>
||强制分页(文章内有效)<br/>
<br/>----换行----<br/>
[br] 换行<br/>
\\ 换行<br/>
<br/>----时间----<br/>
[date] 当前日期<br/>
[time] 当前时间<br/>
[now] 当前具体时间<br/>
[week] 当前星期几<br/>
[month] 当前几月<br/>
[day] 当前几号<br/>
<br/>----图链----<br/>
[img]地址[/img] 显示图片<br/>
[img=图片地址]链接[/img] 带链接的图片<br/>
[url]地址[/url] 显示链接<br/>
[url=地址]文字[/url] 带文字的链接<br/>
<br/>----字体----<br/>
[u]文字[/u] 带下划线的文字<br/>
[b]文字[/b] 加粗的文字<br/>
[i]文字[/i] 倾斜的文字<br/>
[m1]文字[/m1] 向左滚动的文字<br/>
[m2]文字[/m2] 左右滚动的文字<br/>
<br/>----信息----<br/>
[wapname] 网站名称<br/>
[wapurl] 网站地址<br/>
[day=2011-10-1]国庆[/day] 距离国庆还有几天<br/>
[hello] 简单问候语<br/>
[favor] 复杂问候语<br/>
<br/>
<%elseif sip="13" then%>
13、如何添加第三方流量统计代码？<br/>
<br/>
所谓的第三方流量统计就是专门提供站长统计流量的网站，如cnzz，量子，百度等<br/>
统计代码有图片统计代码，js统计代码，wap一般不支持js统计，建议使用图片统计代码<br/>
如何在本系统添加统计代码？<br/>
很简单，只要在“网站配置”-“高级配置”-“网站底部栏目控制”，里面添加即可。这样网站所有的页面都能统计到。<br/>
但要注意，要使用ubb语言哦<br/>
<%elseif sip="14" then%>
14、常用WML标签有哪些？<br/>
<br/>----最常用----<br/>
换行：&lt;br/&gt;<br/>
链接：&lt;a href="网址"&gt;网站名称&lt;/a&gt;<br/>
图片：&lt;img src="图片地址" alt="说明文字"/&gt;<br/>
<br/>----不常用----<br/>
段落：&lt;p&gt;段落&lt;/p&gt;<br/>
&lt;p align="left"&gt;居左段落&lt;/p&gt;<br/>
&lt;p align="center"&gt;居中段落&lt;/p&gt;<br/>
带下划线的文字：&lt;u&gt;文字&lt;/u&gt;<br/>
加粗的文字：&lt;b&gt;文字&lt;/b&gt;<br/>
斜体的文字：&lt;i&gt;文字&lt;/i&gt;<br/>
彩色字体：&lt;font color="red"&gt;文字&lt;/font&gt;<br/>
……<br/>
<br/>
Ps:WML代码可以有机的组合，如图片链接为&lt;a href="网址"&gt;&lt;img src="图片地址" alt="说明文字"/&gt;&lt;/a&gt; ，但不建议新手这么做。<br/>
<br/>
WML标签很严格，编写时请耐心。有些旧的手机不支持部分标签，如斜体，彩色…<br/>
<%elseif sip="15" then%>
15、如何编辑网站底部文字和连接？<br/>
<br/>
网站配置-高级配置-网站底部栏目控制，要使用ubb语言<br/>
<%elseif sip="16" then%>
16、如何随时切换网站风格？<br/>
<br/>
网站配置-网站样式，选择Wap1.0或Wap2.0保存<br/>
<%elseif sip="17" then%>
17、如何过滤用户敏感词？<br/>
<br/>
网站配置-高级配置-敏感词过滤<br/>
注意：用英文逗号分隔，如：色,情<br/>
<%elseif sip="18" then%>
18、如何设置首页文章调用标题字数？<br/>
<br/>
网站配置-高级配置-调用文章标题长度<br/>
注意：留空表示不限制<br/>
<%end if%>
<br/><anchor>返回<prev/></anchor><br/>
<a href="faq.asp?sid=<%=sid%>">[手册中心]</a><br/>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card></wml>
<%call CloseConn%>
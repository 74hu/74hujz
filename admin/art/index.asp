
<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="文章管理">
<p>
<%
response.write"<a href=""wzcl.asp?sid="&sid&""">[文章管理]</a>"&chr(13)
response.write"<br/><a href=""wzclass.asp?sid="&sid&""">[分类管理]</a><br/>"&chr(13)
response.write"<a href=""tjwz.asp?sid="&sid&""">[添加分类]</a>"&chr(13)
response.write"<br/><a href=""adminpl.asp?sid="&sid&""">[文章评论]</a>"&chr(13) 
%>
<br/>----------<br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a><br/>
</p></card></wml>
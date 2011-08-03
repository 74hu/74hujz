
<!--#include file="../mymin.asp"-->
<%IF KEY<>0 and key<>2 then
Call Head()
  Call Error("<card title=""出错""><p>你的权限不足！")
  end if%>
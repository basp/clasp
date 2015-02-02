<!doctype html>	
<html>
<body>
<!--#include file="includes/aspJSON1.17.asp" -->
<!--#include file="includes/http.asp" -->
<%
sub writeln(s)
	response.write "<pre>" & s & "</pre>"
end sub

set data = createobject("Scripting.Dictionary")
data.add "foo", formatdatetime(now, vbshortdate)
data.add "bar", 123

'http://blogs.msdn.com/b/ericlippert/archive/2003/09/15/52996.aspx
dim url: url = "http://localhost/api/json_object.html"
set res = http_get((url), data)
writeln(res.data("foo"))
writeln(url)
%>
</body>
</html>
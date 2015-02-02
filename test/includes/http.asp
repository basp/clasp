<%
'The `http_build_query_array` and `http_build_query` implementations are
'copied from: https://gist.github.com/joseph-montanez/1948929

sub http_build_query_array(key, items, count, out)
	dim result
	for i = 0 to count - 1
		if isarray(items(i)) then
			http_build_query_array (key & "[" & i & "]"), _
				items(i), ubound(items(i)), result
		else
			result = result & server.urlencode (key & "[" & i & "]") & _
				"=" & server.urlencode(items(i))
		end if
		if i + 1 <> count then
			result = result & "&"
		end if
	next
	out = result
end sub

sub http_build_query(keys, items, count, out)
	dim result
	for i = 0 to count - 1
		if isarray(items(i)) then
			http_build_query_array keys(i), _
				items(i), ubound(items(i)), result
		else
			result = result & server.urlencode (keys(i)) & _
				"=" & server.urlencode(items(i))
		end if
		if i + 1 <> count then
			result = result & "&"
		end if
	next
	out = result
end sub

function http_get(url, params)
	dim q

	if not params is nothing then
		http_build_query params.keys, _
			params.items, params.count, q
		url = url & "?" & q
	end if

	set req = server.createobject("MSXML2.ServerXMLHTTP")
	req.open "GET", url, false
	req.send ""

	set res = new aspJSON
	res.loadJSON(req.responseText)
	set http_get = res
end function

function http_post(url, body, params)
	dim q

	if not params is nothing then
		http_build_query params.keys, _
			params.items, params.count, q
		url = url & "?" & q
	end if

	set req = server.createobject("MSXML2.ServerXMLHTTP")
	req.open "POST", url, false
	req.send body

	'TODO: 
	'Optimize this by passing in the URL directly into the `loadJSON` method
	set res = new aspJSON
	res.loadJSON(req.responseText)

	set http_post = res
end function
%>
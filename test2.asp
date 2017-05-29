<%
sub dothis(one, two, name, size)
	
	if 		one = "t" then %>
		<input type="text" name="<%=name%>" value="<%=two%>" size="<%=size%>"><BR>
<%	elseif 	one = "c" then %>
		<input type="checkbox" name="<%=name%>" value="<%=two%>">
<%	elseif 	one = "a" then %>
		<textarea rows="<%=size\2%>" name="<%=name%>" cols="<%=size%>"><%=two%></textarea>
<% 	else %>
	You fucked up - <%=one%>
		
<% end if
	
end sub

	Dim strconn, conn, rst, howmanyfields
	
	strconn="PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE="
	strconn=strconn & server.mappath("svwwtpn.mdb") & ";"
		'strconn=strconn & "Password=whatever;"

	set conn=server.createobject("adodb.connection")	
	conn.open strconn

	set rst	= conn.execute("SELECT * FROM 1")
	howmanyfields=rst.fields.count - 2
	
	for i = 0 to howmanyfields 
	Dim myIdentity
	myIdentity = rst(i).name %>
	<%=Mid(myIdentity, 2)%>:&nbsp;
	<%
	Call dothis(Left(myIdentity, 1), "none", myIdentity, "40")
	
	next	
		
	

'Call dothis("a", "Y", "1", "20")
%>
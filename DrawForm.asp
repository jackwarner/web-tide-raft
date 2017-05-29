<%
sub dothis(one, two, name, size)
	
	if 		one = "t" then %>
		<input type="text" name="<%=name%>" value="<%=two%>" size="<%=size%>"><BR>
<%	elseif 	one = "c" then %>
		<input type="checkbox" name="<%=name%>" value="<%=two%>"
		<% if Len(two) > 0 then
		Response.Write(" checked") 
		end if %>
		>
<%	elseif 	one = "a" or one = "x" then %>
		<textarea rows="<%=size\3%>" name="<%=name%>" cols="<%=size%>"><%=two%></textarea>
<% 	else %>
	You fucked up - <%=one%><% end if
	
end sub%>
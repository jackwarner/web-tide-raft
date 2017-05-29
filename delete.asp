<% Option Explicit %>
<html>
<head>
<title>RAFT ( Delete )</title>
</head>
<body>
<table width="100%">
<tr>
<td><b><font face="Arial,Helvetica"><font size=+1>&nbsp;&nbsp; <a NAME="ADIR">Delete Item(s)</a>
</font></font></b></td>
</tr>
<tr> <td>
<%
Response.Expires = 0
Dim List
List = ""
%>


<%
	Dim Username, myArray, myFormListings, Confirm, WhichTable

	Username	= Session("UName")
	myFormListings 	= Request.Form("DeleteIt")
	myArray		= Split(myFormListings, ",")
	WhichTable	= Request.Form("UDirectory")

	if Request.Querystring("Confirm") = "Yes" then %>
	
		<form method="POST" action="delete.asp">
		<!-- Form Item 1 -->
		<input type="hidden" value="<%=Username%>" name="thisone">
		<!-- Form Item 2 -->
		<input type="hidden" value="<%=WhichTable%>" name="damn">

<% 
	Dim count

	for count = 0 to UBound(myArray) step 1
		Response.Write("<li>" & Mid(myArray(count), Instr(myArray(count), "~")+1) & "</li>") %>
		<input type="hidden" value="<%=Trim(Mid(myArray(count), 1, Instr(myArray(count), "~") - 1))%>" name="<%=Trim(Mid(myArray(count), 1, Instr(myArray(count), "~") - 1))%>">
		
<%  List = List & "<li>" & Mid(myArray(count), Instr(myArray(count), "~")+1) & "</li>"
	
next %>
<BR><hr><input type="hidden" value="<%=List%>" name="Items">
<input TYPE="button" VALUE="  No  " onClick="history.go(-1)">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" value="  Yes  " name="Yep">
</form>
	
<% else %>
<tr>
<td><b><font face="Arial,Helvetica"><font size=+1>&nbsp;&nbsp; <a NAME="ADIR">Success!</a>
</font></font></b></td>
</tr>

<tr>
<td>
<blockquote>
<hr WIDTH="100%"><font face="Arial,Helvetica"><font size=-1>
You succussfully deleted the Item(s) listed below:<BR></font></font><font face="Arial,Helvetica"><font size=-1></font></font>
<%
Dim DBase, i, Directory	
DBase		= Request.Form.Item(1)
Directory 	= Request.Form.Item(2)

Dim strconn, conn, rst, mySQL
	
	strconn="PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE="
	strconn=strconn & server.mappath(DBase) & ".mdb" & ";"
		'strconn=strconn & "Password=whatever;"
	

	for i = 3 to Request.Form.Count - 2
		
		set conn=server.createobject("adodb.connection")
		conn.open strconn
		mySQL="DELETE FROM " & Directory & " WHERE ID = " & Request.Form.Item(i)
		conn.execute(mySQL)
		conn.close
		set conn = nothing
	next
	Response.Write(Request.Form("Items") & "<BR>")
%>
<BR>
<font face="Arial" size="2">[<a href="activedirectory.asp">Go 
Back to Directory Mgmt</a>] [<a href="raft.asp">Go 
Back to RAFT Menu</a>]</font>
<% end if %>
</td></tr>
</table>
</body>
</html>
<% Option Explicit 
   Response.Expires = 0	
%>
<% 	
	Dim UserName, Password, Directory, MaxListing
	
	UserName 			= Session("UName")
	Password 			= Session("UPass")
	
	
	if Len(Request.Form("WhichDir")) > 0 then
		Directory 			= Request.Form("WhichDir")	
		Session("UDir") 	= Directory
	else
		Directory			= Session("UDir")
	end if
	
	Dim strconn, conn, rst, mySQL, rsmax
	
	strconn="PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE="
	strconn=strconn & server.mappath(UserName & ".mdb") & ";"
	'strconn=strconn & "Password=whatever;"
	set conn=server.createobject("adodb.connection")
	conn.open strconn
	
	set rst=conn.execute("select * from " & Directory)
	mySQL="select * from " & Directory & " ORDER BY " & rst(0).name


	set rst=conn.execute(mySQL)
	mySQL = "select MAXLISTING, DIRNAME from ROOT where ID=" & Directory
	set rsmax = conn.execute(mySQL)
	MaxListing 		= rsmax("MAXLISTING") \ 1
%>
<html>
<head>
   <title>Web Tide Studios RAFT</title>
   <meta http-equiv="Expires" CONTENT="0">
  <meta http-equiv="Cache-Control" CONTENT="no-cache">
  <meta http-equiv="Pragma" CONTENT="no-cache">

</head>
<body>
&nbsp;

<div align="center">
  <table BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH="90%" height="525" >
<tr>
<td height="14">
<center>
<b><font size=+1 face="Arial,Helvetica">Manage '<%=rsmax("DIRNAME")%>'</font></font></b></center>
</td>
</tr>
<tr><td>[ <a href="raft.asp"><font size="2" face="Arial">Back To Raft</font></a> ]
</td></tr>
<tr>
<td>
<hr>
  <p style="line-height: 100%; margin-top: 0; margin-bottom: 0"><b><font size=+1 face="Arial,Helvetica">Edit an existing Item</font></b></p>
</td>
</tr>

<tr>
<td>
<blockquote>
<hr WIDTH="100%">
<font face="Arial,Helvetica"><font size=-1>Choose any one of the Items below to edit its properties.</font></font>

<%'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@%>
<form name="Edit" method="POST" action="edit.asp">
<select name="EditDir" size="1">
<% 	Dim i
	i = 0

	do while NOT rst.EOF 
		i = i + 1 %>
		<option value="<%=rst("ID")%>"><%=rst(rst(0).name)%></option>
<% 	rst.MoveNext
	loop 
	
	if i = 0 then %>
		<option value="">No records to edit</option>
	<% else %>
	</select> <input type="submit" name="Submit" value="Edit Item">
<% end if %>
<input type="Hidden" name="UDirectory" value="<%=Directory%>">
</form>
<%'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@%>
</blockquote>
</td>
</tr>

<tr>
<td><b><font face="Arial,Helvetica"><font size=+1>&nbsp;&nbsp;Add a new Item</font></font></b></td>
</tr>

<tr>
<td>
<blockquote>
<hr WIDTH="100%">
<font face="Arial,Helvetica"><font size=-1>
<%if i >= MaxListing then %>
	You have reached the limit of items (<%=MaxListing%>) you may display in this section.  Upgrade your service or delete 
	an item in order to free space.
<% else %></font></font>
<form name="ToAdd" action="add.asp" method="POST">
<input type="Hidden" name="UDirectory" value="<%=Directory%>">
<input type="submit" value="Add a new item to the active directory">
</form>

<% end if %>
</blockquote>
</td>
</tr>

<tr>
<td><b><font face="Arial,Helvetica"><font size=+1>&nbsp;&nbsp; <a NAME="ADIR">Delete an old Item</a>
</font></font></b></td>
</tr>

<tr>
<td>
<blockquote>
<hr WIDTH="100%"><font face="Arial,Helvetica"><font size=-1>
Delete one or more of the items listed below (hold the 'Ctrl' button to select 
multiple items).</font></font>

<%'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@%>

<form name="Delete" method="POST" action="delete.asp?Confirm=Yes">

<select name="DeleteIt" multiple size="<% if i = 0 then
													Response.Write(i+1)
												else
													Response.Write(i)
												end if%>">
<% 	if i = 0 then %>
		<option value="">No records to delete</option>
<%	else

	rst.MoveFirst

	do while NOT rst.EOF  %>
		<option value="<%=rst("ID") & "~" & rst(rst(0).name)%>"><%=rst(rst(0).name)%></option>
<% 		rst.MoveNext
	loop 
	%>
	
</select> <input type="submit" name="Submit" value="Delete Selected">
<%	end if %>
<input type="Hidden" name="UDirectory" value="<%=Directory%>">
</form>
<%'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@%>
</blockquote>
</td>
</tr>
</center>
<tr>
<td>
<p align="right">
</center>
</td>
</tr>
  <center>
<center>
<tr>
<td height="21"></td>
</tr>
</table></center>
  </div>
</center>
</body>
</html>
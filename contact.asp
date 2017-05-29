<% Option Explicit %>
<%
const accessDB 		= "db1.mdb" 

Function DBaseAction(SQLAdd, DataBaseName)
	Dim strconn, conn
	
	strconn="PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE="
	strconn=strconn & server.mappath(DatabaseName) & ";"
		'strconn=strconn & "Password=whatever;"

	set conn=server.createobject("adodb.connection")	
	conn.open strconn

	set DBaseAction	= conn.execute(SQLAdd)
End Function

Dim mySQL, rst
mySQL="select * from ADMIN WHERE USER='" & Session("UName") & "'"
set rst = DBaseAction(mySQL, "db1.mdb")
%>

<html>
<head>
   <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
   <meta name="Author" content="Default">
   <meta name="GENERATOR" content="Microsoft FrontPage 4.0">
   <title>Web Tide Studios RAFT</title>
</head>
<body>
&nbsp;
<center><table BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH="90%" >
<tr>
<td><b><font face="Arial,Helvetica"><font size=+1>&nbsp;&nbsp; Contact
Information</font></font></b></td>
</tr>

<tr>
<td>
<center>
<p></p>
</center>
</td>
</tr>

<tr>
<td>
<blockquote>
<hr WIDTH="100%">
<br><font face="Arial,Helvetica"><font size=-1>View your current contact
information.&nbsp; The information listed is also your billing address
and technical support data.</font></font></blockquote>
</td>
</tr>

<tr>
<td>
<blockquote><b><font face="Arial,Helvetica"><font size=-1>Name:&nbsp; <%=rst("NAME")%></font></font></b>
<br><b><font face="Arial,Helvetica"><font size=-1>Organization:&nbsp; <%=rst("ORG")%>
<hr WIDTH="100%">Phone: <%=rst("PHONE")%></font></font></b>
<br><b><font face="Arial,Helvetica"><font size=-1>Fax: <%=rst("FAX")%></font></font></b>
<br><b><font face="Arial,Helvetica"><font size=-1>Email: <%=rst("EMAIL")%>
<hr WIDTH="100%">Address for billing, contact, and support purposes:</font></font></b>
<p><b><font face="Arial,Helvetica"><font size=-1><%=rst("ORG")%></font></font></b>
<br><b><font face="Arial,Helvetica"><font size=-1>c/o <%=rst("NAME")%></font></font></b>
<br><b><font face="Arial,Helvetica"><font size=-1><%=rst("ADD1")%></font></font></b>
<br><b><font face="Arial,Helvetica"><font size=-1><%=rst("CITY")%>, <%=rst("STATE")%>&nbsp;<%=rst("ZIP")%></font></font></b></blockquote>
</td>
</tr>
</table></center>
<%
	rst.close
  	set rst=nothing
%>
</body>
</html>
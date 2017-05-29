<% Option Explicit %>
<html>
<head>
    <title>Web Tide Studios RAFT</title>
      <meta http-equiv="Expires" CONTENT="0">
  <meta http-equiv="Cache-Control" CONTENT="no-cache">
  <meta http-equiv="Pragma" CONTENT="no-cache">
</head>
<body>
<%
Dim Login, Password
if Len( Request.Form("UName") ) > 0 AND Len( Request.Form("UPass")) > 0 then
	Login	 		= Request.Form("UName") 
	Password 		= Request.Form("UPass")
else
	Login	 		= Session("UName") 
	Password		= Session("UPass")
end if

const accessDB = "db1.mdb" 	
	Dim strconn, FlagComm, FlagDir, FlagMail
	strconn="PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE="
	strconn=strconn & server.mappath(accessDB) & ";"
		'strconn=strconn & "Password=whatever;"
	Dim mySQL, rst
	mySQL		= "SELECT USER, PASS, NAME, ACTIVEDIR, ECOMMERCE, MAILLIST from ADMIN WHERE PASS='" & Password & "' AND USER='" & Login & "'"
  dim conntemp, rstemp, howmanyfields
   set conntemp=server.createobject("adodb.connection")
   conntemp.open strconn
   set rst=conntemp.execute(mySQL)

if rst.EOF then %>
	<p align="center" style="text-indent: 0; line-height: 100%; margin: 0"><b><font face="Arial" size="5" color="#FF0000"><i>Sorry!&nbsp;
</i></font></b></p>
<hr>
<ul>
  <li>
    <p align="left" style="text-indent: 0; line-height: 100%; margin: 0"><font face="Arial" size="2"><b>The
    User Name/Password combination you entered is invalid.&nbsp; </b></font></li>
  <li>
    <p align="left" style="text-indent: 0; line-height: 100%; margin: 0"><font face="Arial" size="2"><b>Remember
    that these fields are <i>CaSe SenSitive</i>!&nbsp; </b></font></li>
  <li>
    <p align="left" style="text-indent: 0; line-height: 100%; margin: 0"><font face="Arial" size="2"><b>Click
    the 'Back' button on your browser to try again. </b></font></li>
</ul>
<% else 
Session("UName") = Login
Session("UPass") 	= Password
FlagComm 	= rst("ECOMMERCE")
FlagDir 	= rst("ACTIVEDIR")
FlagMail 	= rst("MAILLIST")
%>
<center><table BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH="90%" >
<tr>
<td colspan="2">
<center>

<p>

<b><font face="Arial,Helvetica"><font size=-1></font></font></b></p>
</center>
<p>&nbsp;</td>
</tr>

<tr>
<td><b><font size="3" face="Arial,Helvetica">Welcome, <%=rst("NAME")%>, </b></font></font></font></b><font face="Arial,Helvetica" size="2">to the
<b>R</b>eal-life <b>A</b>pplication <b>F</b>or <b>T</b>echnology web maintenance system, an internet service providing instant access to your personal
account and contact information with Web Tide Studios.&nbsp;&nbsp;</font>
</td>
<td>
<p align="center"><img border="0" src="raft_logo_s.gif"></td>
</tr>

<tr>
<td colspan="2">
<hr WIDTH="100%"></td>
</tr>

<tr>
<td colspan="2"><b><font face="Arial,Helvetica"><font size=+1>&nbsp;&nbsp; Contact
Information</font></font></b></td>
</tr>

<tr>
<td colspan="2">
<blockquote>
<hr WIDTH="100%">
<br><font face="Arial,Helvetica"><font size=-1><a href="contact.asp">View your current contact
information.&nbsp; The information listed is also your billing address
and technical support data.</a></font></font><font face="Arial,Helvetica"><font size=-2></font></font>
<div align=right>
<p><font face="Arial,Helvetica"><font size=-2>[ Contact us to change the
information in this section ]</font></font></div>
</blockquote>
</td>
</tr>

<tr>
<td colspan="2"><b><font face="Arial,Helvetica"><font size=+1>&nbsp;&nbsp; Account
Summary</font></font></b></td>
</tr>

<tr>
<td colspan="2">
<blockquote>
<hr WIDTH="100%">
<br><font face="Arial,Helvetica"><font size=-1><a href="account.asp">View your currently selected
billing options as well as the due date and dollar amount of your next
payment.</a></font></font><font face="Arial,Helvetica"><font size=-1></font></font>
<div align=right>
<p><font face="Arial,Helvetica"><font size=-2>[ Contact us to change the
information in this section ]</font></font></div>
</blockquote>
</td>
</tr>

<tr>
<td colspan="2"><b><font face="Arial,Helvetica"><font size=+1>&nbsp;&nbsp; <a NAME="ADIR">Active Directories</a>

<% if FlagDir = "Yes" then
	Response.Write("En")
   else
	Response.Write("Dis")
   end if 

   rst.close
   set rst=nothing
   conntemp.close
   set conntemp=nothing

%>abled</font></font></b></td>
</tr>

<tr>
<td colspan="2">
<blockquote>
<hr WIDTH="100%"><font face="Arial,Helvetica"><font size=-1>
<% 	if FlagDir = "Yes" then
	Dim conn
	strconn="PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE="
	strconn=strconn & server.mappath(Login & ".mdb") & ";"
		'strconn=strconn & "Password=whatever;"

	mySQL="select * from ROOT ORDER BY DIRNAME"
	
	set conn=server.createobject("adodb.connection")
	conn.open strconn
	set rst=conn.execute(mySQL)
%>
<form name="Choice" method="POST" action="activedirectory.asp">
<select name="WhichDir" size="1">
<% 	rst.MoveFirst
	do while NOT rst.EOF %>
		<option value="<%=rst("ID")%>"><%=rst("DIRNAME")%></option>
<% 	rst.MoveNext
	loop %>
</select> <input type="submit" name="Submit" value="Manage Directory">
</form>
<% 
   rst.close
   set rst=nothing
   conn.close
   set conn=nothing

else %>

RAFT may also
be extended to enable real-time updating of products, services, or events
on your professionally designed and hosted web site.&nbsp; Please contact
us to enable active directories.</font></font><font face="Arial,Helvetica"><font size=-1></font></font>

<% end if %>
<div align=right>
<p><font face="Arial,Helvetica"><font size=-2>[ Contact us to change the
information in this section ]</font></font></div>
</blockquote>
</td>
</tr>
<%
	if FlagComm = "Yes" then
%>
<tr>
<td colspan="2"><b><font face="Arial,Helvetica"><font size=+1>&nbsp;&nbsp; E-Commerce
Enabled</font></font></b></td>
</tr>
<tr>
<td colspan="2">
<blockquote>
<hr WIDTH="100%"><font face="Arial,Helvetica"><font size=-1>RAFT may be
extended to enable secure, cost-effective E-Commerce solutions for the
successful business wishing to tap into the potential market the Internet
has to offer.&nbsp; Please contact us to configure your E-Commerce solution.</font></font><font face="Arial,Helvetica"><font size=-1></font></font>
<div align=right>
<p><font face="Arial,Helvetica"><font size=-2>[ Contact us to change the
information in this section ]</font></font></div>
</blockquote>
</td>
</tr>
<% end if %>
<%
	if Len(FlagMail) > 0 then
%>
<tr>
<td colspan="2"><b><font face="Arial,Helvetica"><font size=+1>&nbsp;&nbsp; Mailing List Management</font></font></b></td>
</tr>
<tr>
<td colspan="2">
<blockquote>
<hr WIDTH="100%"><font face="Arial,Helvetica"><font size=-1><a href="http://www.webtidestudios.com/email">Click here</a> to log onto
 the mailing list management site.  Use the same user name/password combination you entered for RAFT, and choose the 'Compose' option.
 Type '<%=FlagMail%>' in the 'To:' field to send out a mass email.  Remember, you can change your settings at any time by choosing
 'List Administration' from the drop down box in the upper right hand area of the IMail system.
</font></font><font face="Arial,Helvetica"><font size=-1></font></font>
<div align=right>
<p><font face="Arial,Helvetica"><font size=-2>[ Contact us to change the
information in this section ]</font></font></div>
</blockquote>
</td>
</tr>
<% end if %>

<tr>
<td colspan="2">
<center>[ <font face="Arial" size="2"><a href="index.asp?Action=Logout">Log Out</a></font> ]</center>
</td>
</tr>

<tr>
<td colspan="2"></td>
</tr>
</table></center>
<% end if %>
</body>
</html>
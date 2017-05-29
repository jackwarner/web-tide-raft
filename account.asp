<% Option Explicit %>
<%
const accessDB 		= "db1.mdb" 
const DELIMIT		= ";"

Dim Login, ToDisplay
	Dim UserName, Password	
	UserName = Session("UName")
	Password = Session("UPass")
	
Dim mySQL
Dim rst

Dim TotalCharge
TotalCharge = 0

Dim ProdLims(3, 1)
' Is the same as rates table, save dbase connection
' First define max number of products
ProdLims(0, 0) = 10
ProdLims(1, 0) = 25
ProdLims(2, 0) = 50
ProdLims(3, 0) = 100
' Now define factor to multiply
ProdLims(0, 1) = 1
ProdLims(1, 1) = 1.2
ProdLims(2, 1) = 1.5
ProdLims(3, 1) = 2

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GET ADMIN INFORMATION
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	mySQL="select * from ADMIN WHERE USER='" & UserName & "'"
	
Function DBaseAction(SQLAdd, DataBaseName)
	Dim strconn, conn
	
	strconn="PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE="
	strconn=strconn & server.mappath(DatabaseName) & ";"
		'strconn=strconn & "Password=whatever;"

	set conn=server.createobject("adodb.connection")	
	conn.open strconn

	set DBaseAction	= conn.execute(SQLAdd)
End Function

	
	set rst=DBaseAction(mySQL, "db1.mdb")
	
	Dim BillInterval, DueDate, FlagDir, BaseEmails, ExtraEmails, FlagComm

	BillInterval 	= rst("BILLME")
	DueDate 		= rst("PAYMENTDUE")
	FlagDir		= rst("ACTIVEDIR")
	FlagComm		= rst("ECOMMERCE")
	BaseEmails		= rst("EMAILACCT")  \ 1
	ExtraEmails	= rst("EXTRAEMAIL") \ 1

	rst.close
  	set rst=nothing

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GET BILLING CYCLE INFORMATION, ACCORDING TO INTERVAL		
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	mySQL="select * from BILLCYCLE WHERE CYCLE='" & BillInterval & "'"
	set rst	= DBaseAction(mySQL, "db1.mdb")


	Dim PerYear, BaseHosting, BaseDir, BaseEmailCharge

	PerYear		= rst("FACTOR")
	BaseHosting	= rst("HOSTING")
	BaseDir		= rst("DIRECTORY") \ 1
	BaseEmailCharge = rst("EMAILCHARGE")	

	rst.close
  	set rst=nothing

%>
	
<html>
<head>
   <title>Web Tide Studios RAFT</title>
</head>
<body>
&nbsp;
<center><table BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH="90%" >
<tr>
<td><b><font face="Arial,Helvetica"><font size=+1>&nbsp;&nbsp; Account
Summary</font></font></b></td>
</tr>

<tr>
<td>

</td>
</tr>

<tr>
<td>
<blockquote>
<hr WIDTH="100%">
<br><font face="Arial,Helvetica"><font size=-1>View your currently selected
billing options as well as the due date and dollar amount of your next
payment.</font></font></blockquote>
</td>
</tr>

<tr>
<td>
<blockquote><b><font face="Arial,Helvetica"><font size=-1>The following Web Services
for your account are billed on a(n) <%=BillInterval%> basis:</font></font></b><b><font face="Arial,Helvetica"><font size=-1></font></font></b>
<p><b><font face="Arial,Helvetica"><font size=-1>Web Hosting: $<%=BaseHosting%>.00</font></font></b>
<br><b><font face="Arial,Helvetica"><font size=-1>Included Email Accounts:
<%=BaseEmails%> @ $0.00 Each</font></font></b>
<br><b><font face="Arial,Helvetica"><font size=-1>Additional Email Accounts:
<%=ExtraEmails%> @ $<%=BaseEmailCharge%>.00 Each</font></font></b>
<br><b><font face="Arial,Helvetica"><font size=-1>Active Directories Enabled:&nbsp;&nbsp;<%=FlagDir%></font></font></b>
<blockquote>
<blockquote>

<%
if FlagDir = "Yes" then
		mySQL="select * from ROOT"
		set rst=DBaseAction(mySQL, UserName & ".mdb")

 rst.MoveFirst
	Dim i, ADCharge
	ADCharge = 0.0
	i = 0
	do while NOT rst.EOF
		i = i + 1
		Response.Write("<font face=Arial,Helvetica><font size=-1><li>A maximum of " & rst("MAXLISTING") & " listings for <u>" & rst("DIRNAME") & "</u> @ $")
		Dim j 
		j = 0

		do while j < 4

			if (rst("MAXLISTING") \ 1) <= ProdLims(j, 0) then
				ADCharge = BaseDir * ProdLims(j, 1)   				
				Response.Write(ADCharge & ".00</li></font></font>")
				TotalCharge = TotalCharge + ADCharge
				j = 4
			end if
			j = j + 1
		loop
		rst.MoveNext
	loop

	rst.close
  	set rst=nothing

end if
	%>
</blockquote>
</blockquote>
<b><font face="Arial,Helvetica"><font size=-1>E-Commerce:&nbsp;<%=FlagComm%></font></font></b><b><font face="Arial,Helvetica"><font size=-1></font></font></b>
<p><b><font face="Arial,Helvetica"><font size=-1>Total Due: $<%= (TotalCharge + BaseHosting + (ExtraEmails * BaseEmailCharge)) %>.00 by <%=DueDate%> (
<% Dim DueWhen
DueWhen = DateDiff("d", Now, CDate(DueDate))
if DueWhen < 0 then %>
<font color="#FF0000"><%=DueWhen* -1%>&nbsp;days ago</font>
<% elseif DueWhen = 0 then %>
today
<% else %>
<%=DueWhen%> days from now
<% end if %> ).

</font></font></b>
<p><b><font face="Arial,Helvetica"><font size=-1>Annual Total: $<%= (TotalCharge + BaseHosting + (ExtraEmails * BaseEmailCharge)) * PerYear %>.00 </font></font></b>
<p><b><font face="Arial,Helvetica"><font size=-1>Please may check payable
to:</font></font></b>
<blockquote>
  <p style="margin-top: 0; margin-bottom: 0"><b>
  <font size=-1 face="Arial,Helvetica">Web Tide Studios</font></b>
<br><b><font face="Arial,Helvetica" size="-1">209 Barger Street</font></b></p>
  <p style="margin-top: 0; margin-bottom: 0"><b>
  <font face="Arial,Helvetica" size="-1">Blacksburg, VA 24060</font></b></p>
  </blockquote>
</blockquote>
</td>
</tr>
</table></center>

</body>
</html>
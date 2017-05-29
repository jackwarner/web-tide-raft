<% option explicit %>
<html>
<!--#include file="DrawForm.asp"-->
<!--#include file="ConnectDBase.asp"-->
<% 
Dim rst, mySQL
mySQL = "SELECT * FROM " & Request.Form("UDirectory") & " WHERE ID=" & Request.Form("EditDir") \ 1
rst = DBaseAction(mySQL, Session("UName") & ".mdb")
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>RAFT ( Edit )</title>
</head>

<body>

<div align="center">
  <center>
  <form method="POST" action="exesql.asp">
  <table border="0" width="705" cellspacing="0" cellpadding="0" height="178">
    <tr>
      <td colspan="2" height="43" width="703"><b><font size="+1" face="Arial,Helvetica">Edit
        an Item</font></b>
        <hr>
      </td>
    </tr>
     <%
    Dim strconn, conn, rst2, howmanyfields, i
	
	strconn="PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE="
	strconn=strconn & server.mappath("svwwtpn.mdb") & ";"
		'strconn=strconn & "Password=whatever;"

	set conn=server.createobject("adodb.connection")	
	conn.open strconn
	Dim myDir 
	myDir = "SELECT * FROM " & Request.Form("UDirectory")

	set rst2	= conn.execute(myDir)
	
	howmanyfields=rst2.fields.count - 2
	
	for i = 0 to howmanyfields 
	Dim myIdentity
	myIdentity = rst(i).name  %>
	 <tr>
      <td width="207" valign="top"><font face="Arial" size="2"><%=Replace(Mid(myIdentity, 2), "_", " ")%>:</font></td>
	  <td width="494"><font face="Arial" size="2">
		<%	Call dothis(Left(myIdentity, 1), rst(myIdentity), myIdentity, "40") %></font></td>
    </tr>
	<% next %>    
    <tr>
      <td height="27" colspan="2" width="703">
        <p align="center"><input type="submit" value="Submit" name="B1"> <input type="reset" value="Reset" name="B2"> <input TYPE="button" VALUE="Back" onClick="history.go(-1)"></td>
    </tr>
  </table>

<input type="Hidden" name="UDirectory" value="<%=Request.Form("UDirectory")%>">
<input type="Hidden" name="UItem" value="<%=Request.Form("EditDir")%>">
<input type="Hidden" name="Confirmation" value="edit_ok.asp">
  </form>
  </center>
</div>


</body>

</html>
<% option explicit %>
<html>
<head>
   <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
   <meta name="Author" content="Default">
   <meta name="GENERATOR" content="Microsoft FrontPage 5.0">
   <title>RAFT</title>
</head>
<body>

<form action="raft.asp" method="POST">
<center>
<p>
<div align="center">
  <center>
  <p></p>
  <form action="raft.asp" method="post">
  <div align="center">
    <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="52%" id="AutoNumber1">
    <tr>
      <td width="50%" rowspan="2"><img border="0" src="raft_logo.gif"></td>
      <td width="19%"><b><font face="Arial" size="1">User Name</font></b></td>
      <td width="32%"><input size="20" name="UName" style="float: left"></td>
    </tr>
    <tr>
      <td width="19%"><b><font face="Arial" size="1">Password</font></b></td>
      <td width="32%"><input type="password" size="20" name="UPass"></td>
    </tr>
    <tr>
      <td width="100%" colspan="3">
<hr></td>
    </tr>
    <tr>
      <td width="50%">
      <p align="center"><b>
<font face="Arial,Helvetica" size="2" color="#0000FF"><i>&quot;Putting the Web where it belongs -- in
Your hands&quot;</i></font></b></td>
      <td width="25%">&nbsp;</td>
      <td width="25%">
<input type="submit" name="Click" value=" Login ">
<input type="reset" value="Reset" name="B2"></td>
    </tr>
  </table>
<% if Request.Querystring("Action") = "Logout" then
	Session.Abandon %>
	<hr width="52%"><i><b><font face="Arial" size="2" color="#0033CC">Thank you for using RAFT!</font></b></i>   
	<% end if %>
    </center>
  </div>

  </center>
</div>

  </form>
</body>
</html>
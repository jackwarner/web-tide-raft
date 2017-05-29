<% 	option explicit 
	Dim ToDo, mySQL
	ToDo		= Request.Form("Confirmation")
Function DBaseAction(SQLAdd, DataBaseName)
	Dim strconn, conn
	
	strconn="PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE="
	strconn=strconn & server.mappath(DatabaseName) & ";"
		'strconn=strconn & "Password=whatever;"

	set conn=server.createobject("adodb.connection")	
	conn.open strconn

	set DBaseAction	= conn.execute(SQLAdd)
End Function


Dim temp, i, j, k
if 		ToDo 		= "add_ok.asp" 	then
mySQL = "INSERT INTO " & Request.Form("UDirectory") 
	mySQL = mySQL & " ("
	for i = 1 to Request.Form.Count - 3
		mySQL = mySQL & Request.Form.Key(i) 
		if i <> Request.Form.Count - 3 then
			mySQL = mySQL & ", "
		end if
	next
	mySQL = mySQL & ") VALUES ("
	for j = 1 to Request.Form.Count - 3
		mySQL = mySQL & "'" & Replace(Replace(Request.Form.Item(j), Chr(34), "'"), "'", "''") & "'"
		if j <> Request.Form.Count - 3 then
			mySQL = mySQL & ", "
		end if
	next
	mySQL = mySQL & ")"
	
elseif ToDo		= "edit_ok.asp" 	then

mySQL 	= 	"UPDATE " & Request.Form("UDirectory") 
		mySQL = mySQL & " SET "
		for k = 1 to Request.Form.Count - 4
			mySQL = mySQL & Request.Form.Key(k) & " = '" & Replace(Replace(Request.Form.Item(k), Chr(34), "'"), "'", "''") & "'"
			if k <> Request.Form.Count - 4 then
				mySQL = mySQL & ", "
			end if
		next 
		mySQL = mySQL & "WHERE ID=" & Request.Form("UItem") \ 1

end if


Dim rst
rst = DBaseAction(mySQL, Session("UName") & ".mdb")

Response.Redirect(ToDo)
%>
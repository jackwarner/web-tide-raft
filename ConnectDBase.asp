<% Function DBaseAction(SQLAdd, DataBaseName)
	Dim strconn, conn
	
	strconn="PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE="
	strconn=strconn & server.mappath(DatabaseName) & ";"
		'strconn=strconn & "Password=whatever;"

	set conn=server.createobject("adodb.connection")	
	conn.open strconn

	set DBaseAction	= conn.execute(SQLAdd)
End Function %>
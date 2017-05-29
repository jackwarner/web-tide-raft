<% OPTION EXPLICIT %>
<% 
sub query2table(inputquery, inputDSN)
   dim conntemp, rstemp, howmanyfields
   set conntemp=server.createobject("adodb.connection")
   conntemp.open inputDSN
   set rstemp=conntemp.execute(inputquery)
   howmanyfields=rstemp.fields.count -1%>
   <table border=1><tr>
   <% 'Put Headings On The Table of Field Names
Dim i, thisvalue
   for i=0 to howmanyfields %>
             <td><b><%=rstemp(i).name%></B>&nbsp;</TD>
   <% next %>
   </tr>
   <% ' Now lets grab all the records
   do while not rstemp.eof %>
      <tr>
      <% for i = 0 to howmanyfields
         thisvalue=rstemp(i)
         If isnull(thisvalue) then
            thisvalue="&nbsp;"
         end if%>
             <td valign=top><%=thisvalue%>&nbsp;</td>
      <% next %>
      </tr>
      <%rstemp.movenext
   loop%>
   </table>
   <%
   rstemp.close
   set rstemp=nothing
   conntemp.close
   set conntemp=nothing
end sub%>



<%	const accessDB = "db1.mdb" 	
	Dim strconn
	strconn="PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE="
	strconn=strconn & server.mappath(accessDB) & ";"
		'strconn=strconn & "Password=whatever;"
Dim mySQL
mySQL="select * from ADMIN"

call query2table(mySQL,strconn)

%>
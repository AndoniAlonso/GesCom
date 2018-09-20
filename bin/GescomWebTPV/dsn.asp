<%
Option Explicit

'Open your connection
Dim objConn
Set objConn = server.CreateObject("ADODB.Connection")
	
'to do: change the database path below to your own database location
objConn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=GestionComercial;Data Source=localhost"
%>
<!-- #include file="connect.inc" -->
<%
var Records=Server.CreateObject("ADODB.Recordset")
Records.ActiveConnection=Connect
Records.CursorType = 0
Records.CursorLocation = 2
Records.LockType = 3

var Records1=Server.CreateObject("ADODB.Recordset")
Records1.ActiveConnection=Connect
Records1.CursorType = 0
Records1.CursorLocation = 2
Records1.LockType = 3
%>
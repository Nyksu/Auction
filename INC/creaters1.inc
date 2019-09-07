<%
function CreateRecordSet1(){
	var recs1=Server.CreateObject("ADODB.Recordset")
	recs1.ActiveConnection=Connect
	recs1.CursorType = 0
	recs1.CursorLocation = 2
	recs1.LockType = 3
	return recs1
}
%>
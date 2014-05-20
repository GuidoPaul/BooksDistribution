<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Buffer = True %>
<%
	response.Clear() '修改乱码, 关键的一句
	dim conn
	dim connstr
	connstr="driver={sql server};database=books;server=127.0.0.1;uid=sa;pwd=;"
	set conn = Server.CreateObject("ADODB.Connection")
	conn.open connstr
%>

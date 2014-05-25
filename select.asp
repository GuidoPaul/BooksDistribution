<!--#include file="conn.asp"-->
<html>
	<head>
		<meta content="text/html;charset=gb2312" http-equiv="Content-Type">
		<title>Select</title>
	</head>
	<body>
	<%
		Set rs = server.CreateObject("ADODB.Recordset")
		S_Magor = Session("S_Magor")
		S_Class = Session("S_Class")
		S_ClassID = Session("S_ClassID")

		sql = "select * from Book"
		rs.open sql, conn

		Set rs2 = server.CreateObject("ADODB.Recordset")
		sql2 = "Select C_Number from Class where C_Magor = '" & S_Magor & "'"
		rs2.open sql2, conn
		C_Number = rs2("C_Number")
		Session("C_Number") = C_Number


		Set rs3 = server.CreateObject("ADODB.Recordset")
		sql3 = "Select * from Book where B_ClassID = '" & Trim(S_ClassID) & "'"
		rs3.open sql3, conn

		with response
			if rs3.EOF then
				.write "现在数据库为空!"
			else
				.write "<table border=1 cellspace=0 cellpadding=5 align=center>" &_
					"<tr height=12>" &_
						"<td width=50><strong>Magor</strong></td>" &_
						"<td width=180><strong>" & S_Magor & "</strong></td>" &_
						"<td width=50><strong>Class</strong></td>" &_
						"<td width=50><strong>" & S_Class & "</strong></td>" &_
					"</tr>" &_
					"<tr height=12>" &_
						"<td><strong>BookISBN</strong></td>" &_
						"<td><strong>BookName</strong></td>" &_
						"<td><strong>TakeNumber</strong></td>" &_
						"<td><strong>StockNumber</strong></td>" &_
					"</tr>"
			end if
			do until rs3.EOF
				.write "<tr height=12>" &_
					"<td>" & rs3("B_ISBN") & "</td>" &_
					"<td>" & rs3("B_Name") & "</td>" &_
					"<td>" & C_Number & "</td>" &_
					"<td>" & rs3("B_StockNumber") & "</td>"
			rs3.movenext
			loop
			.write "</table>"
		end with

		rs.close
		set rs = nothing
		conn.close
		set conn = nothing

	%>
		<form action="ok.asp" method="post">
			<p style="text-align:center;">
				<input type="submit" value="OK" >
				<input type="reset" value="Reset" >
			</p>
		</form>
	</body>
</html>

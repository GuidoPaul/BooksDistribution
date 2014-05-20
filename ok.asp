<!--#include file="conn.asp"-->
<html>
	<head>
		<meta content="text/html;charset=UTF-8" http-equiv="Content-Type">
		<title>OK</title>
	</head>
	<body>
	<%
		Set rs = server.CreateObject("ADODB.Recordset")
		S_ClassID = Session("S_ClassID")
		C_Number = Session("C_Number")
		C_Number = cint(C_Number)
		sql = "Select * from Book where B_ClassID = '" & Trim(S_ClassID) & "'"
		rs.open sql, conn, 1, 3

		''response.write C_Number

		do
			rs("B_StockNumber") = rs("B_StockNumber") - C_Number
			rs("B_TakenNumber") = rs("B_TakenNumber") + C_Number
		rs.movenext
		loop until rs.EOF

		''rs.update '奇怪, 为什么不能写这句?'
		rs.close
		set rs = nothing
		conn.close
		set conn = nothing
	%>
		<p style="text-align:center">
			操作完成，数据库已修改
		</p>
		<form action="index.asp" method="post">
			<p style="text-align:center;">
				<input type="submit" value="返回首页" >
			</p>
		</form>
	</body>
</html>

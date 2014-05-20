<!--#include file = "conn.asp"-->
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
		<title>For</title>
	</head>

	<body>
		<%
			dim S_Magor, S_Class_str, Password, p

			S_Magor = request.form("S_Magor")
			S_Class_str = request.form("S_Class")
			Password = request.form("Password")

			set rs = server.createobject("ADODB.Recordset")
			sql = "Select M_Password from Manager"
			rs.open sql, conn, 1, 1

			set rs2 = server.createobject("ADODB.Recordset")
			sqlSS = "Select * from Student"
			rs2.open sqlSS, conn, 1, 1

			Function check()
				dim flag, S_Class
				S_Class = cint(S_Class_str) '转为整数, 若为空值转换时会出错, 故写在函数中, 放在正确的地方调用
				flag = false
				''response.write S_Magor & " " & S_Class & "<br />"
				do
					''response.write rs2.Fields("S_Magor") & " " & rs2.Fields("S_Class") & "<br />"
					if rs2.Fields("S_Magor") = S_Magor and rs2.Fields("S_Class") = S_Class then
						flag = true
						Session("S_ClassID") = rs2.Fields("S_ClassID")
					end if
					rs2.movenext
				loop until rs2.EOF or flag = true
				check = flag
			end Function

			if S_Magor = "" or S_Class_str = "" then
			%>
				<script language="vbscript">
					alert("学生专业或班级不能为空")
					history.back()
				</script>
			<%
			elseif Password = "" then
			%>
				<script language="vbscript">
					alert("管理员密码不能为空")
					history.back()
				</script>
			<%
			elseif check = false then
			%>
				<script language="vbscript">
					alert("学生专业或班级输入错误")
					history.back()
				</script>
			<%
			else
				if rs.Fields("M_Password") = Password then
					Session("S_Magor") = S_Magor
					Session("S_Class") = S_Class_str
					response.redirect "select.asp"
				else
				%>
					<script language="vbscript">
						alert("管理员密码错误！")
						history.back()
					</script>
				<%
				end if
			end if
		%>
	</body>
</html>


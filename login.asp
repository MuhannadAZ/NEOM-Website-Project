<%
Dim Conn
Dim SQLQuery
Dim rs
Dim UserName
Dim Password

UserName = request.form("uname")
Password = request.form("psw")
RememberMe = request.form("rememberme")

if UserName <> "" or Password <> "" then
	set Conn=server.createobject("ADODB.Connection")
	set rs = Server.CreateObject("ADODB.Recordset")
	connStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\inetpub\wwwroot\NEOM\Database1.mdb; User ID=; Password=;"

	Conn.open  connStr
	
	SQLQuery = "select * from Database1 where email = '"&UserName&"' AND psw = '"&Password&"'"

	set rs=Conn.execute(SQLQuery)
	if rs.BOF and rs.EOF then
        Response.redirect "loginError.html"
	else 
        Response.Cookies("UserName")=UserName
        Response.Cookies("Password")=Password
        Response.Redirect "home.html"
		Conn.Close
		rs.close   
		set rs = nothing
		set Conn = nothing
	end if
else
    response.write("<script language=""javascript"">alert('Invalid Username or password');</script>")
end if
%>
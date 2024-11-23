<%
Dim Conn
Dim SQLQuery
Dim rs
Dim UserName
Dim Password
Dim pswRepeat

UserName = request.form("email")
Password = request.form("psw")
pswRepeat = Request.Form("psw-repeat")

if UserName <> "" or Password <> "" then
	set Conn=server.createobject("ADODB.Connection")
	set rs = Server.CreateObject("ADODB.Recordset")
	connStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\inetpub\wwwroot\NEOM\Database1.mdb; User ID=; Password=;"

	Conn.open  connStr
	
	SQLQuery = "select * from Database1 where email = '"&UserName&"'"

	set rs=Conn.execute(SQLQuery)
	if rs.BOF and rs.EOF then
        dbInsert = "INSERT INTO Database1 (email,psw, pswRepeat) VALUES ('"& UserName &"','"& Password &"','"& pswRepeat &"')"
        conn.Execute(dbInsert)
        response.redirect "login.html"
	else 
        response.redirect "signupError.html"
		Conn.Close
		rs.close   
		set rs = nothing
		set Conn = nothing
	end if
else
    response.write("<script language=""javascript"">alert('Invalid Username or password');</script>")
end if
%>
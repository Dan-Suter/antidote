	<%
	Session.Contents.RemoveAll()
	session.Abandon
	response.redirect("/default.asp")
	%>
	
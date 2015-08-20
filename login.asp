<!--#include virtual="/header.htm" -->
<!--#include virtual="/functions.asp" -->
<!--#include virtual="/connection.asp" -->
	<%
	sLoginAttempt=""
	if not request.querystring("a")="" then
		'autoLogin by person to serve another person.
		'x=rwe("here")
		sSQL= "CALL `antidote`.`People_login_by_Auto`('"&request.querystring("a")&"');"		
		x=openRS(sSQL)
		if not rsTemp.eof then
			session("id_people")=rsTemp("id_people")
			session("email")=rsTemp("email")
			session("password")=rsTemp("password")
			session("name")=rsTemp("name")
			session("image_path")=rsTemp("image_path")
			session("uid_people")=rsTemp("uid_people")
			session("about_me")=rsTemp("about_me")
			session("can_authorize")=rsTemp("can_authorize")
			if not request("r")="" then
				response.redirect(request("r"))
			else
				response.redirect("/my_home.asp")
			end if			
			x=closeRS()
		end if
		x=closeRS()



	end if
	if not request.form("username")="" then
		sSQL= "CALL `antidote`.`People_login`('"&request.form("username")&"','"&request.form("password")&"');"		
		x=openRS(sSQL)
		if not rsTemp.eof then
			session("id_people")=rsTemp("id_people")
			session("email")=rsTemp("email")
			session("password")=rsTemp("password")
			session("name")=rsTemp("name")
			session("image_path")=rsTemp("image_path")
			session("uid_people")=rsTemp("uid_people")
			session("about_me")=rsTemp("about_me")
			session("can_authorize")=rsTemp("can_authorize")
			response.redirect("/my_home.asp")
			x=closeRS()
		end if
		x=closeRS()
		sLoginAttempt="fail"
	end if
	sErr="" 
	%>
<form action="/login.asp" method="post" name="frmLogin" id="frmLogin">
<div class="row row-centered" style="padding:10px 10px 10px 10px;">
	<div class="col-xs-12" style="text-align:center;">
		<h3>Please enter your email and password.</h3>
	</div>
</div>
<div class="row row-centered" style="padding:10px 10px 10px 10px;">
	<div class="col-xs-5" style="text-align:right;">
		User name
	</div>
	<div class="col-xs-7">
		<input name="UserName" type="text" size="24" value="">
	</div>
</div>
<div class="row row-centered" style="padding:10px 10px 10px 10px;">
	<div class="col-xs-5" style="text-align:right;">
		Password
	</div>
	<div class="col-xs-7">
		<input name="PassWord" type="password" value="" size="24">
	</div>
</div>
<div class="row row-centered" style="padding:10px 10px 10px 10px;">
	<div class="col-xs-12" style="text-align:center;">
		<input name="Login" type="submit" value="Login to website" size="20" style="padding:10px;">
	</div>
</div>
<%if sLoginAttempt="fail" then %>
<div class="row row-centered" style="padding:10px 10px 10px 10px;">
	<div class="col-sd-12" style="text-align:center;">
		<p class="error">* Incorrect User Name / password combination.</p>
	</div>
</div>
<%end if%>
</form>

<!--#include virtual="/footer.asp" -->
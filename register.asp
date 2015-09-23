<!--#include virtual="/header.htm" -->
<!--#include virtual="/functions.asp" -->
<!--#include virtual="/connection.asp" -->
<script>
  window.fbAsyncInit = function() {
    FB.init({
      appId      : '1575727666027370',
      xfbml      : true,
      version    : 'v2.3'
    });
  };

  (function(d, s, id){
     var js, fjs = d.getElementsByTagName(s)[0];
     if (d.getElementById(id)) {return;}
     js = d.createElement(s); js.id = id;
     js.src = "//connect.facebook.net/en_US/sdk.js";
     fjs.parentNode.insertBefore(js, fjs);
   }(document, 'script', 'facebook-jssdk'));
</script>
<script>
AIzaSyAtJ6uukYKLx-vRZNJioqvDbj-W7zTdlcw
</script>

	<%
	sLoginAttempt=""
	if not request.form("username")="" then
		sSQL= "CALL `antidote`.`People_check`('"&request.form("email")&"')"
		'x=rwb(sSQL)
		x=openRS(sSQL)
		if not rsTemp.eof then
		'user already exists
			sregister="fail"
		else
			'add User
			x=closeRS()
			sSQL= "CALL `antidote`.`People_Add`('"&request.form("username")&"','"&request.form("email")&"','"&request.form("password")&"')"
			x=openRS(sSQL)
			x=closeRS()
			sSQL= "CALL `antidote`.`People_login`('"&request.form("email")&"','"&request.form("password")&"')"
			x=openRS(sSQL)
			session("id_people")=rsTemp("id_people")
			session("name")=rsTemp("name")
			session("email")=rsTemp("email")
			session("password")=rsTemp("password")
			x=closeRS()
			response.redirect("/my_home.asp")
			x=closeRS()
		end if
		x=closeRS()
		sLoginAttempt="fail"
	end if
	sErr="" %>
<form action="/register.asp" method="post" name="frmLogin" id="frmLogin">

<div class="login_form">
    <input name="Login_fb" style="width: 100%;" class="signup_facebook login_button" type="button" value="Register using Facebook" onclick="document.location.href='/fb_login.asp'">   
</div>

<div class="login_form">
    <span class="signup_google login_button">
        Register using Google
    </span>
</div>
<div class="login_form">
    <h3>Or</h3>
</div>
<div class="row row-centered" style="padding:10px 10px 10px 10px;">
	<div class="col-xs-5 col-sm-5" style="text-align:right;">
		Your Name
	</div>
	<div class="col-xs-7 col-sm-7">
		<input name="UserName" type="text" size="24" value="<%=request.form("UserName")%>">
	</div>
</div>
<div class="row row-centered" style="padding:10px 10px 10px 10px;">
	<div class="col-xs-5 col-sm-5" style="text-align:right;">
		Email Address
	</div>
	<div class="col-xs-7 col-sm-7">
		<input name="email" type="text" size="24" value="<%=request.form("email")%>">
	</div>
</div>
<div class="row row-centered" style="padding:10px 10px 10px 10px;">
	<div class="col-xs-5 col-sm-5" style="text-align:right;">
		Password
	</div>
	<div class="col-xs-7 col-sm-7">
		<input name="PassWord" type="password" value="<%=request.form("UserName")%>" size="24">
	</div>
</div>
<div class="row row-centered" style="padding:10px 10px 10px 10px;">
	<div class="col-xs-12" style="text-align:center;">
		<input name="termsandconditions" type="checkbox" style="padding:10px;">
		 I agree to the <a href="/help/terms_and_conditions.asp" id="terms">Terms and Conditions.</a> of the site.
	</div>
</div>
<div class="row row-centered" style="padding:10px 10px 10px 10px;">
	<div class="col-xs-12 login_form" style="">
		<input name="Login" class="login_button" type="submit" value="Sign Up" style=""> 
	</div>
</div>
<div
  class="fb-like"
  data-share="true"
  data-width="450"
  data-show-faces="true">
</div>
<%if sregister="fail" then %>
<div class="row row-centered" style="padding:10px 10px 10px 10px;">
	<div class="col-xs-1"  style="text-align:center;">
		<p class="error">* Email address already exists. click <a href="/login.asp">here</a> to login</p>
	</div>
</div>
<%end if%>
</form>

<!--#include virtual="/footer.asp" -->
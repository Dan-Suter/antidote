<!--#include virtual="/header.htm" -->
<!--#include virtual="/functions.asp" -->
<!--#include virtual="/connection.asp" -->
<!--#include file="fb_app.asp" -->
<!--#include file="fb_graph_api_app.asp" -->
<script language="javascript" runat="server" src="json2.asp">

</script>

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
			'session("can_authorize")=rsTemp("can_authorize")
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
    if not request.form("Login_fb")="" then
        dim user
        
        set user = new fb_user
        user.token = cookie("token")
        user.LoadMe
        
        WriteBR "Token: " & cookie("token")
        WriteBR "Expires: " & cookie("expires")
        WriteBR "URL: " & user.graph_url
        WriteBR ""
        WriteBR "Json: "
        WriteBR user.json_string
        WriteBR ""
        WriteBR "-----------------------------------------"
        WriteBR "RESULTS"
        WriteBR "-----------------------------------------"
        
        WriteBR "First Name: " & user.first_name
        WriteBR "Last Name: " & user.last_name
        WriteBR "Email: " & user.email
        WriteBR "Picture: " & user.m_id
        DrawPicture "https://graph.facebook.com/" & user.m_id & "/picture"
    end if
	sErr="" 
    
    
    
function WriteBR( str )
    response.write str
    response.write "<br/>"
end function

function DrawPicture( str )
    response.write "<img src="
    response.write str
    response.write ">"
end function

	%>

    
<form action="/login.asp" method="post" name="frmLogin" id="frmLogin">

<div class="login_form">
    <input name="Login_fb" class="signup_facebook login_button" type="submit" value="Sign in with Facebook">   
</div>

<div class="login_form">
    <span class="signup_google login_button">
        Sign in with Google
    </span>
</div>
<div class="login_form">
    <h3>Or</h3>
</div>

<div class="login_form">
    <input name="UserName" type="text" size="30" value="" placeholder="Username or email">
</div>
<div class="login_form">
    <input name="PassWord" type="password" value="" size="30" placeholder="PassWord">
</div>

<div class="login_form">
    <input name="Login" type="submit" value="Login">   
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
<!--#include virtual="/header.htm" -->
<!--#include virtual="/functions.asp" -->
<!--#include virtual="/connection.asp" -->
<!--#include virtual="/fb_app.asp" -->
<!--#include virtual="/fb_graph_api_app.asp" -->

<script language="javascript" runat="server" src="json2.asp">

</script>

<%
    'if not request.form("Login_fb")="" then
        dim user
        
        set user = new fb_user
        user.token = cookie("token")
        user.LoadMe
        'so first thing to check is does the users email address already exist?'
        'if so then then user is simply trying to login not register'
        'use facebook to authenticate login process'
        sSQL="People_Check ('"&user.email&"')"
        openRS(sSQL)
        'x=rwe(sSQL)
        if rsTemp.eof then
        	if not len(user.email)=0 then
	        	' No User so lets add this person to our DB'
	        	sSQL= "CALL `antidote`.`People_Add_FB`('"&user.first_name &" " & user.last_name &"','"&user.email&"','"&user.m_id&"')"
	        	'x=rwb(sSQL)
	        	x=openRS(sSQL)
	        	x=closeRS()
	        	sSQL= "Select * from people where fb_id='"&user.m_id&"';"
				x=openRS(sSQL)
				if not rsTemp.eof then
					session("id_people")=rsTemp("id_people")
					session("email")=rsTemp("email")
					session("password")=rsTemp("password")
					session("name")=rsTemp("name")
					'session("image_path")=rsTemp("image_path")
					id_people=rsTemp("uid_people")
					session("uid_people")=rsTemp("uid_people")
					session("about_me")=rsTemp("about_me")
					session("can_authorize")=rsTemp("can_authorize")
					x=closeRS()
				end if
				x=closeRS()
				sLoginAttempt="fail"
	        	'asuming no error get uid_people'
        		pic="http://graph.facebook.com/" & user.m_id & "/picture?width=150&height=200"
        		'get pictutre and sizes'
        		sSavePath=sFilePath&"\images\people\original\"&id_people&".jpg"
        		'/images/people/med/b555d96a.jpg'
        		'x=rwb(pic&":"&sSavePath)
        		'x=rwb(pic&":"&sSavePath)
        		sOriginalPath=sSavePath
        		call Get_Image(pic,sSavePath)
			 	pic="http://graph.facebook.com/" & user.m_id & "/picture?width=600&height=800"
			 	sSavePath=sFilePath&"\images\people\xlarge\"&id_people&".jpg"
			 	call Resize_Image(sOriginalPath,sSavePath,600,800)
			 	sSavePath=sFilePath&"\images\people\large\"&id_people&".jpg"
			 	call Resize_Image(sOriginalPath,sSavePath,375,450)
			 	sSavePath=sFilePath&"\images\people\med\"&id_people&".jpg"
			 	call Resize_Image(sOriginalPath,sSavePath,188,225)
			 	sSavePath=sFilePath&"\images\people\small\"&id_people&".jpg"
				call Resize_Image(sOriginalPath,sSavePath,94,112)
				sSavePath=sFilePath&"\images\people\thumb\"&id_people&".jpg"
				call Resize_Image(sOriginalPath,sSavePath,46,62)
				sSavePath=sFilePath&"\images\people\xsthumb\"&id_people&".jpg"
				call Resize_Image(sOriginalPath,sSavePath,23,31)
				'send email with login details and welcome pack'
				'sweet we have logged in all the photos are loaded, update the img_path'
				sSQL="update antidote.people set Image_path=concat('/images/people/med/','"&id_people&"','.jpg')  where fb_id='"&id_people&" and id_people>0;"
				openRS(sSQL)
				closeRS

			 	response.redirect("/my_home.asp")
        	else
        		'error using facebook redirect to register page with error msg'
        		x=rda("/login.asp")
        	end if
        	
        else
				 sSQL= "Select * from people where fb_id='"&user.m_id&"';"
				x=openRS(sSQL)
				if not rsTemp.eof then
					session("id_people")=rsTemp("id_people")
					session("email")=rsTemp("email")
					session("password")=rsTemp("password")
					session("name")=rsTemp("name")
					'session("image_path")=rsTemp("image_path")
					session("uid_people")=rsTemp("uid_people")
					session("about_me")=rsTemp("about_me")
					session("can_authorize")=rsTemp("can_authorize")
					x=closeRS()
					x=rda("/my_home.asp")  
				else
					x=closeRS()   
					x=rda("/login.asp") 
				end if
				
				  	
        end if

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
        DrawPicture "https://graph.facebook.com/" & user.m_id & "/picture?type=large"
    'end if
	sErr="" 
    
    
    
function WriteBR( str )
    response.write str
    response.write "<br/>"
end function

function DrawPicture( str )
    response.write "<img src="""
    response.write str
    response.write """>"
end function

	%>

    
<form action="/fb_login.asp" method="post" name="frmLogin" id="frmLogin">

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
    <input name="Login" class="login_button"  type="submit" value="Login">   
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
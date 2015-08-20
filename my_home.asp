<!--#include virtual="/header.htm" -->
<!--#include virtual="/functions.asp" -->
<!--#include virtual="/connection.asp" -->
<!--#include virtual="/security.asp" -->
<style>
.thumb-wrapper {position:relative;}
.thumb-wrapper span {position:absolute;top: 0px;left: 0px;width: 100%;height: 100%;z-index: 100;background: transparent url(/images/whitecam.png) no-repeat;}
</style>
<%
		sSQL= "CALL `antidote`.`People_login_By_ID`('"&session("id_people")&"');"
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
		end if
		x=closeRS()
		sLoginAttempt="fail"
	%>
		<form id="myform" name="myform" action="/admin/ajax/save.asp">
			<input type="hidden" name="t" value="people">
			<input type="hidden" id="uid" name="uid" value="<%=session("uid_people")%>">
			<input type="hidden" id="file_name" name="file_name" value="<%=session("uid_people")%>.jpg">
			
			<div class="row row-centered">
				<div class="col-md-12">
					<h3>My Profile</h3>
				</div>
			</div>
 			<div class="row row-centered">
		        <div class="col-md-4 indent10" >
		          My Name
		        </div>
		        <div class="col-md-8">
				<input type="text" size="30" name="name" id="name" value="<%=session("name")%>">  
		        </div>
		    </div>
 			<div class="row row-centered">
		        <div class="col-md-4 indent10" >
		          Email
		        </div>
		        <div class="col-md-8">
		        	<input type="text" size="30" name="email" id="email" value="<%=session("email")%>">      
		        </div>
		    </div>
 			<div class="row row-centered">
		        <div class="col-md-4 indent10" >
		          Password
		        </div>
		        <div class="col-md-8">
		        	<input type="password" size="30" name="password" id="password" value="<%=session("password")%>">           
		        </div>
		    </div>
 			<div class="row row-centered">
		        <div class="col-md-4 indent10" >
		          My Picture
		        </div>
		        <div class="col-md-8" style="">

				<%if isnull(session("image_path")) or session("image_path")="" then%>
				<div>
				<div id="uploader">
					<p>Your browser doesn't have Flash, Silverlight or HTML5 support.</p>
				</div>
				<br />
				<%
				else%>
				<div class="thumb-wrapper">
				<img id="update_img" src="<%=session("image_path")%>" alt="<%=session("name")%>" />
				<span id="updatePhoto" onclick="updatePhoto();"></span>
				<pre id="log" style="height: 300px; overflow: auto;display:none;"></pre>
				</div>
					
				<%end if
				%>           
		        </div>
		    </div>

		    <div class="row row-centered">
		        <div class="col-md-4 indent10" >
		         My Story
		        </div>
		        <div class="col-md-8">
		          <textarea id="about_me" name="about_me" class="editor"><%=session("about_me")%></textarea>
		        </div>
		    </div>
		  
		    <div class="row row-centered" style="text-align:center">
		    	<div class="col-md-12">
		            <button type="button"  class="button" onclick="Save_Form()">Save Changes</button>
		      </div>
		    </div>
			<div class="row row-centered">
				<div class="col-md-2">
					<h3>My Recipes</h3>
				</div>
		    	<div class="col-md-19">
		           <a class="button icon add" href="/add_Recipe.asp">Add a New Recipe..</a>
		      </div>
			</div>

		    <%'check to see if person has any Recipes
			sSQL= "CALL `antidote`.`Recipes_By_Person`('"&session("id_people")&"');"
					'x=rwb(sSQL)
					x=openRS(sSQL)
					do until rsTemp.eof%>

					<div class="row row-centered">
				        <div class="col-md-4 indent10" >
				          <img src="<%=rsTemp("image")%>" alt="<%=rsTemp("name")%>">
				        </div>
				        <div class="col-md-6">
							<%=rsTemp("name")%>  added on (<%=rsTemp("added")%>) 
				        </div>
				        <div class="col-md-2">
				         	<a class="button icon edit" href="add_Recipe.asp?id=<%=rsTemp("id_recipe")%>">Edit</a>
				        </div>
				    </div>	
				    <%					
				    rsTemp.movenext
				    loop
					x=closeRS()

					x=closeRS()	%>
		   </form>
			<div id="popup" style="left: 701.5px; position: absolute; top: 106px; z-index: 9999; opacity: 1; display: none; background:#fff;">
	        <span class="button b-close"><span>X</span></span>
					<div id="uploader">
						<p>Your browser doesn't have Flash, Silverlight or HTML5 support.</p>
					</div>
	    </div>
<!--#include virtual="/footer.asp" -->
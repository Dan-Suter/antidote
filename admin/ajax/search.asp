<!--#include virtual="/security.asp" -->
<!--#include virtual="/functions.asp" -->
<!--#include virtual="/connection.asp" -->
<%
'get URL and get Recipe food item values
if request("t")="p" then
	
	sSQL=""
	sSQL="SELECT id_people,    email,    password,    name,    image_path,    uid_people,    about_me,    can_authorize,auto_login FROM `antidote`.`people` where name like '"&request("v")&"%';"
	x=openRS(sSQL)
	irow=0
	do until rsTemp.eof
	if irow mod 2=1 then strClass="light_blue_row" else strClass="white_row"
	irow=irow+1
	name=rsTemp("name")
	image_path="/images/people/small/"&rsTemp("uid_people")&".jpg"
	about_me=rsTemp("about_me")
	name=rsTemp("name")
	about_me=rsTemp("about_me")
	id_people=rsTemp("id_people")
	auto_login=rsTemp("auto_login")
	%>
	<div id="htm<%=id_people%>" class="row row-centered " style="height:135px;overflow:hidden;">
    <div class="col-xs-2" style="">
    	
 		<%if session("can_authorize") then%>   	
    	<a href="/login.asp?a=<%=auto_login%>&r=/recipes.asp">
    	<%end if%>
			<%if not image_path="" and not isnull(image_path) then%>
				<img src="<%=image_path%>" alt="<%=name&"'s picture"%>">
				<%
			else
				%>
				<img src="/people/images/blank-face-icon.png" alt="<%=name%>'s picture">
			<%end if%>
 			<%if session("can_authorize") then%>   	
    	</a>
    	<%end if%>
			
			 
			<button id="spn<%=id_people%>" class="button icon arrowdown" onclick="showMore(<%=id_people%>)">Show more.</button>          
    </div>
    <div class="col-xs-10">
      <b><%=name%>.</b> <%=about_me%>
    </div>
  </div>
  <hr>
<%
	rsTemp.movenexT
loop
x=closeRS()
end if
%>
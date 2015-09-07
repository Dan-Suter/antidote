<!--#include virtual="/header.htm" -->
<!-- #include virtual="/connection.asp"-->
<!-- #include virtual="/functions.asp"-->
<script type="text/javascript" src="https://www.google.com/jsapi"></script>

<%if session("name")="" then %>
<div class="foods container-fluid" >
	<div class="row">
		<div class="col-sm-12">
		  	<h1>Anitdote Recipes.</h1>
		 </div>
	</div>
<%else %>	
<div class="foods container-fluid" >
	<div class="row">
		<div class="col-sm-12">
		  	<h1>Hey <%=session("name") %>, what would you like to eat?</h1>
		 </div>
	</div>
<%end if %>	

<%if session("id_person")=id_person or session("can_authorize") then%>
	<div class="row">
			<div class="col-sm-12 col-xs-12">
				<h3>Yay you are someone who can <a class="button icon edit" href="/add_recipe.asp">Add A Recipe</a></h3>
			</div>
	</div>
<%end if%>

<% 

'check to see if id is stated?
'************************************************
'update added grouping to the SQL , also added a column to determine sorting order.
'************************************************
bAdmin=0
if session("can_authorize") then bAdmin=1
sSQL = "Call Get_recipes ("&bAdmin&");"
x=openRS(sSQL)
%>

<%
	irow=0
	sFoodType=""
	do until rsTemp.eof
		if irow mod 2=1 then strClass="light_blue_row" else strClass="white_row"
		irow=irow+1
		id_recipe=rsTemp("id_recipe")
		id_person=rsTemp("id_person")
		name=rsTemp("name")
		image=rsTemp("image")
		how_to_make=rsTemp("how_to_make")
		id_type=rsTemp("id_type")
		servings=rsTemp("servings")
		uid_recipe=rsTemp("uid_recipe")
		%>
		<%if sFoodType<>rsTemp("group_name") then%>
		<div class="row">
			<div class="col-xs-12">
			  	
			  	<h2>
			  		<%=rsTemp("group_name")%>s
			  	</h2>
			  	
			 </div>
		</div>
		<%end if
	  	if sFoodType<>rsTemp("group_name") then sFoodType=rsTemp("group_name") 
	  	%>
		<div class="row">
			<div class="col-sm-10 col-xs-10">
				<h3 style="margin-top:0px;"><a href="/recipe.asp?r=<%=id_recipe%>"><%=name%></a></h3>
			</div>
			<%if session("can_authorize") then%>
			<div class="col-sm-2 col-xs-2">
				<a class="button icon edit" href="/add_recipe.asp?id=<%=id_recipe%>">Edit</a>
			</div>
			<%end if%>
		</div>
		<div class="row" id="htm<%=id_recipe%>" style="height:188px;overflow:hidden;">
			<div class="col-md-4 col-sm-4 col-xs-12">
				<a href="/recipe.asp?r=<%=id_recipe%>"><img src="<%=rsTemp("image")%>" alt="<%=rsTemp("name")%>"></a>
			</div>
			<div class="col-md-4 col-sm-4 col-xs-12">
				<div class="row">		        
			        <div id="ingredients_list" class="col-xs-12">
			        	<ul>
			        	<%'check to see what ingredients are already added?'
			        	x=openRSA("call Recipes_By_ID ("&id_recipe&")")
			        	if rsTempA.eof then
			        		x=rw("No Ingredients")
			        	else
			        		do until rsTempA.eof
			        			x=rw("<li>"&rsTempA("qty_grams")&" grams of <a href=""/food.asp?f="&rsTempA("id_food")&""">"&rsTempA("name")&"</a></li>")
			        			rsTempA.movenext
			        		loop
			        	end if
			        	x=closeRSA()
			        	%>
			        	</ul>
			        	<div><b>Serves <%=servings%> people</b></div>
			        	<a href="/recipe.asp?r=<%=id_recipe%>">View the nutritional content.</a>
			        </div>
			    </div>
			</div>		
			<div class="col-md-4 col-sm-4 col-xs-12">
				<div class="row">
			        <div class="col-sm-12 col-xs-12">
			         	<table>
			         	<%'Graph added 7/09/2015 Dan.
			         	sSQL=""
			         	sSQL="Call Get_recipe_cache ("&rsTemp("id_recipe")&");"
			         	
			         	x=openRSA(sSQL)
			         	'x=rwe(sSQL)
			         	do until rsTempA.eof
			         		width=cint(rsTempA("RDI"))
			         		if width>100 then width=100

			         		x=rw("<tr><td align=""right"" class=""small-graph-name""><a href=""/vitamin.asp?v="&rsTempA("id_vitamin")&""">"&rsTempA("name")&"</a></td><td style=""width:100%""><a href=""/vitamin.asp?v="&rsTempA("id_vitamin")&"""><div class=""small-graph-line""  style=""width:"&width&"%;background-color:"&rsTempA("color")&""">&nbsp;</div></a></td></tr>")
			         		rsTempA.movenext
			         	loop
			         	x=closeRSA()
			         	%>
			         </table>
			        </div>
				</div>
			</div>
		</div>
	<%
	rsTemp.movenext
loop
x=closeRS()%>
	</div>
</div>
<div id="spacer" style="margin-top:20px;"></div>

<!--#include virtual="/footer.asp" -->
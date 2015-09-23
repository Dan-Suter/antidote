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

<%if session("can_authorize") then%>
	<div class="row">
			<div class="col-sm-12 col-xs-12">
				<h3><a class="button icon edit" href="/add_recipe.asp">Add A Recipe</a></h3>
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
id_person=session("id_people")
if id_person="" then id_person=0
sSQL = "Call Get_recipes ("&bAdmin&","&id_person&");"
'x=rwe(sSQL)
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
		uid_recipe=rsTemp("uid_people")
		person_name=rsTemp("person_name")



		%>
		<%if sFoodType<>rsTemp("group_name") then%>
		<div class="row">
			<div class="col-xs-12">
			  	<h2 style="">
			  		<%=rsTemp("group_name")%>s
			  	</h2>
			  	
			 </div>
		</div>
		<%end if
	  	if sFoodType<>rsTemp("group_name") then sFoodType=rsTemp("group_name") 
	  	%>
		<div class="row">
			<div class="col-sm-10 col-xs-10">
				<h3 style=""><a href="/recipe.asp?r=<%=id_recipe%>"><%=name%></a></h3>
			</div>
			<div class="col-sm-2 col-xs-2">
				<%if session("can_authorize") then%>
					<a class="button icon edit" href="/add_recipe.asp?id=<%=id_recipe%>">Edit</a>
				<%end if%>
			</div>
		</div>
		<div class="row" id="htm<%=id_recipe%>" style="height:auto;overflow:hidden;">
			<div class="col-md-4 col-sm-4 col-xs-12">
				<a href="/recipe.asp?r=<%=id_recipe%>"><img src="<%=rsTemp("image")%>" alt="<%=rsTemp("name")%>"></a>
			</div>
			<div class="col-md-4 col-sm-4 col-xs-12">
				<div class="row">		        
			        <div id="ingredients_list" class="col-xs-12">
			        	<ul class="no-indent">
			        	<%'check to see what ingredients are already added?'
			        	x=openRSA("call Recipes_By_ID ("&id_recipe&")")
			        	if rsTempA.eof then
			        		x=rw("No Ingredients")
			        	else
			        		icount=0
			        		do until rsTempA.eof
			        			icount=icount+1
			        			x=rw("<li style=""white-space:nowrap"">"&rsTempA("qty_grams")&" grams of <a href=""/food.asp?f="&rsTempA("id_food")&""">"&rsTempA("name")&"</a></li>")
			        			rsTempA.movenext
			        			if icount=6 then exit do
			        		loop
			        	end if
			        	x=closeRSA()
			        	%>
			        	<li><b>Serves <%=servings%></b></li>
			        	</ul>
        				<div class="row">
			        		<div class="col-md-5 col-sm-5 col-xs-12">
			        		<button type="button" class="btn btn-default btn-sm"  onclick="Add_favourite(<%=rsTemp("id_recipe")%>)">
			        			<%icon="glyphicon-star-empty"
			        			if not isnull(rsTemp("id_people_favourite")) then
			        				icon="glyphicon-star" 
			        			end if
			        			%>
									  <span id="favourite<%=rsTemp("id_recipe")%>" class="glyphicon <%=icon%>" aria-hidden="true"></span> Favourite
										<span class="result"></span>
									</button>
									</div>
									<div class="col-md-7 col-sm-7 col-xs-12">
										<img id="recipes-person-img" src="/images/people/xsthumb/<%=rsTemp("uid_people")%>.jpg" alt="<%=rsTemp("person_name")%>"><b>By <%=rsTemp("person_name")%><i class="icon-large icon-search"></i></b>
									</div>
			        	</div>
			
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

			         		x=rw("<tr><td align=""right"" class=""small-graph-name""><a href=""/vitamin.asp?v="&rsTempA("id_vitamin")&""">"&rsTempA("name")&"</a></td><td style=""width:100%""><a title="""&rsTempA("RDI")&"% of your Recommended Daily Intake"" href=""/vitamin.asp?v="&rsTempA("id_vitamin")&"""><div class=""small-graph-line""  style=""width:"&width&"%;background-color:"&rsTempA("color")&""">&nbsp;</div></a></td></tr>")
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
<script>
$("#favourite").hover()
function Add_favourite(idR)
{
	if ($("#favourite"+idR ).attr("class")=="glyphicon glyphicon-star-empty")
	{
		$("#favourite"+idR ).addClass("glyphicon-star");
		$("#favourite"+idR ).removeClass("glyphicon-star-empty");
		$.get( "/admin/ajax/add_favourite_recipe.asp?r="+idR+"");
		}
		else 
			{
		$("#favourite"+idR ).addClass("glyphicon-star-empty");
		$("#favourite"+idR ).removeClass("glyphicon-star");
		$.get( "/admin/ajax/add_favourite_recipe.asp?r="+idR+"&d=1");
				}

}
</script>
<!--#include virtual="/footer.asp" -->
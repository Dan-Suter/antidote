<!--#include virtual="/header.htm" -->
<!--#include virtual="/functions.asp" -->
<!--#include virtual="/connection.asp" -->
<!--#include virtual="/security.asp" -->

			<%'check to see if user has an active open recipe
			if request("id")<>"" then
				sSQL="SELECT * FROM antidote.recipes where id_recipe="&request("id")&";"
			else
				sSQL="SELECT * FROM antidote.recipes where id_person="&session("id_people")&" and temp=1;"
			'x=rwe(sSQL)
			end if
			x=openRS(sSQL)
			if rsTemp.eof then
				x=closeRS()
				sSQL="call Insert_Temp_Recipe ("&session("id_people")&")"
				x=openRS(sSQL)
				x=closeRS()
				sSQL="SELECT * FROM antidote.recipes where id_person="&session("id_people")&" and temp=1;"
				x=openRS(sSQL)
			end if
			id_recipe=rsTemp("id_recipe")
			id_person=rsTemp("id_person")
			name=rsTemp("name")
			image=rsTemp("image")
			how_to_make=rsTemp("how_to_make")
			id_type=rsTemp("id_type")
			uid_recipe=rsTemp("uid_recipe")
			bshow_on_web=rsTemp("show_on_web")
			bauthorized=rsTemp("authorized")
			servings=rsTemp("servings")
			brief=rsTemp("brief")
			bauthCheck=""
			if bauthorized then bauthCheck="Checked"
			if bshow_on_web then bwebCheck="Checked"
			x=closeRS()
			'x=rwe(sSQL)
			%>
			<form  id="myform" name="myform" action="/admin/ajax/save.asp">
			<input type="hidden" name="t" value="recipe">
			<input type="hidden" name="uid" value="<%=uid_recipe%>">
			<input type="hidden" name="file_name" id="file_name" value="<%=uid_recipe%>.jpg">
			<input type="hidden" name="folder_name" id="folder_name" value="recipe">
			<input type="hidden" name="idr" id="idr" value="<%=id_recipe%>">
			<div class="row row-centered">
				<div class="col-xs-12">
					<h3>Adding Recipe</h3>
				</div>
			</div>
 			<div class="row row-centered vcenter">
		        <div class="col-xs-3 indent10" >
		          Name it
		        </div>
		        <div class="col-xs-5">
					<input type="text" size="40" maxlength="100" name="name" value="<%=name%>">        
		        </div>
		        <div class="col-xs-4">
					<a class="button icon arrowright" href="/recipe.asp?r=<%=id_recipe%>">Preview on web<a>
		        </div>
		    </div>
			<div class="row row-centered">
		        <div class="col-xs-3 indent10" >
		          recipe Picture
		        </div>
		        <div class="col-xs-9">
							<div class="thumb-wrapper">
							<img id="update_img" src="<%=image%>" alt="<%=name%>" />
							<span id="updatePhoto" class="photo_holder_recipe" onclick="updatePhoto();"></span>
							<pre id="log" style="height: 300px; overflow: auto;display:none;"></pre>
							</div>          
		        </div>
		    </div>
 				<div class="row row-centered">
		        <div class="col-xs-3 indent10" >
		          Recipe Type
		        </div>
		        <div class="col-xs-9">
						<select name="type"><option value="0">Select Type</option>
						<%'enumerate foods list
						x=openRS("Select * from Recipe_types order by name;")
						do until rsTemp.eof%>
							<option value="<%=rsTemp(0)%>"<%if id_type=rsTemp(0) then x=rw(" selected") %>><%=rsTemp(1)%></option>
						<%rsTemp.movenext
						loop
						x=closeRS()
						%>
						</select>    
		          for no of people
						<input type="text" name="servings" id="servings" value="<%=servings%>" size="2">    
		        </div>
		    </div>

 				<div class="row row-centered">
		        <div class="col-xs-3 indent10" >
		         	Add serving sizes
		        </div>
		        <div class="col-xs-9">
							<select name="portion_size" id="portion_size"><option value="0">Select</option>
							<%'enumerate foods list
								x=openRS("SELECT id_portion_size,name FROM antidote.portion_sizes order by id_portion_size;")
								do until rsTemp.eof
									%>
									<option value="<%=rsTemp(0)%>"<%if id_type=rsTemp(0) then x=rw("selected") %>><%=rsTemp(1)%></option>
									<%
									rsTemp.movenext
								loop
							x=closeRS()
							%>
							</select>
							is <input type="text" size="2" maxsize="5" id="contribuition_amount" name="contribuition_amount" value="5.00">
							<input type="button" id="contribution" class="button" value="Add Contribution" onclick="add_contribution(<%=id_recipe%>)">  
						</div>
					</div>
					<div>
					<div class="row row-centered">
		        <div class="col-xs-3 indent10" >
		          Serving sizes
		        </div>
		        <div id="contribution_list" class="col-xs-9">
			        	<%'check to see what ingredients are already added?'
								sSQL="Call get_recipe_with_contribution ("&id_recipe&")"
								x=openRSA(sSQL)
								prices=""
								icount=0
			        	if rsTempA.eof then
			        		x=rw("<b><i>Add some pricing options.</i></b>")
			        	else
									do until rsTempA.eof
										icount=icount+1
										x=rw("<div id=""portion_size"&rsTempA("id_recipe_contribution")&""">"&rsTempA("portion_name")&" $"&rsTempA("amount_currency")&" <button class=""button danger icon remove""  onclick=""Delete_Contribution("&rsTempA("id_recipe_contribution")&"); return false;"">Remove</button></div>")
										rsTempA.movenext
									loop
				        end if
			        	%>
		        </div>
		    </div>

 			<div class="row row-centered">
		        <div class="col-xs-3 indent10" >
		          Add Ingredients
		        </div>
		        <div class="col-xs-9">
					<input type="text" size="6" maxsize="5" id="food_amount" name="food_amount" value="100">
					grams of 
					<select id="food_add" name="food_add">
						<%'enumerate foods list
						x=openRS("Select * from food order by name;")
						do until rsTemp.eof%>
							<option value="<%=rsTemp(0)%>"><%=rsTemp(1)%></option>
						<%rsTemp.movenext
						loop
						x=closeRS()
						%>
					</select>
					<input type="button" id="add_ingredient" class="button" value="Add Ingredient" onclick="Add_Ingredient(<%=id_recipe%>)">
					
		        </div>
		    </div>
 			<div class="row row-centered">
		        <div class="col-xs-3 indent10" >
		          Ingredients
		        </div>
		        <div id="ingredients_list" class="col-xs-9">
		        	<%'check to see what ingredients are already added?'
		        	x=openRS("call Recipes_By_ID ("&id_recipe&")")
		        	if rsTemp.eof then
		        		x=rw("<b><i>Add some ingredients to your menu by using the add ingredient button above....</i></b>")
		        	else
		        		do until rsTemp.eof
		        			x=rw("<div id=""ingredient"&rsTemp("id_recipe_food")&""">"&rsTemp("qty_grams")&" grams of <a href=""/food.asp?f="&rsTemp("id_food")&""">"&rsTemp("name")&"</a> <button class=""button danger icon remove""  onclick=""Delete_Ingredient("&rsTemp("id_recipe_food")&"); return false;"">Remove item</button></div>")
		        			rsTemp.movenext
		        		loop
		        	end if
		        	%>
		        </div>
		    </div>
			<div class="row row-centered">
		        <div class="col-xs-3 indent10" style="margin-top:20px;" >
		         	Brief Description
		        </div>
		        <div class="col-xs-9">
						<textarea id="brief" name="brief" cols="100" rows="2" placeholder="Add a Brief Description of the recipe here.  This text will show on Menu page."><%=brief%></textarea>
		        </div>
		    </div>
 			<div class="row row-centered">
		        <div class="col-xs-3 indent10" style="margin-top:20px;" >
		         	How to Make it
		        </div>
		        <div class="col-xs-9">
					<textarea id="makeit" name="makeit" class="editor"><%=how_to_make%></textarea>
		        </div>
		    </div>

 		    <div class="row " style="">
 		    	<div class="col-xs-4"></div>
		    	<div class="col-xs-8">
		    		<input type="checkbox" name="show_on_web" id="show_on_web" checked="<%=bwebCheck%>"> <label for="show_on_web">Show this recipe to the live website.</label></br>
		           <%if session("can_authorize")=true then%>
		           <input type="checkbox" name="authorized" id="authorized" checked="<%=bauthCheck%>"> <label for="authorized">Authorize this recipe.</label>
		           <%end if%>
		      </div>
		    </div>			
		    <div class="row row-centered" style="text-align:center">
				<div class="col-xs-5">
		      	</div>
		    	<div class="col-xs-4 text-left">
		           <button  class="button" onclick="Save_Form(); return false;">Save Changes</button>
		      	</div>
		    	<div class="col-xs-3 text-left">
		           <button  class="button danger icon remove" onclick="Delete_Record('recipe','<%=uid_recipe%>','/recipes.asp'); return false;">Delete Record</button>
		      </div>
		    </div>
		</form>
			<div id="popup" style="left: 701.5px; position: absolute; top: 106px; z-index: 9999; opacity: 1; display: none; background:#fff;">
		        <span class="button b-close"><span>X</span></span>
						<div id="uploader">
							<p>Your browser doesn't have Flash, Silverlight or HTML5 support.</p>
						</div>
		    </div> 
			</form>
			<!--#include virtual="/footer.asp" -->
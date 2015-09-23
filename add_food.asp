<!--#include virtual="/header.htm" -->
<!--#include virtual="/functions.asp" -->
<!--#include virtual="/connection.asp" -->
<!--#include virtual="/security.asp" -->
			<%
			
			'check to see if user chosen to add a recipe of edit one?
			if request("id")<>"" then
				id_food=request("id")
			else
				'bad ID report to user.
			
				sSQL="call Insert_Temp_Food ("&session("id_people")&")"
				
				x=openRS(sSQL)
				x=closeRS()
				'get the ID of the food generated.
				sSQL="SELECT max(id_food) as id_food_max FROM antidote.food where id_person_add="&session("id_people")&";"
				x=openRS(sSQL)
				id_food=rsTemp(0)
				x=closeRS()
			end if
			
			sSQL="SELECT * FROM antidote.food where id_food="&id_food&";"
			
			x=openRS(sSQL)
			if rsTemp.eof and len(request("id"))>0 then
				'bad ID report to user.
				x=rwinfo("This is a bad request food ID. Please try again")
			end if
			id_food=rsTemp("id_food")
			id_person=rsTemp("id_person_add")
			uid_food=rsTemp("uid_food")
			name=rsTemp("name")
			Intro=rsTemp("Intro")
			Image_path=rsTemp("Image_path")
			Description=rsTemp("Description")
			default_unit=rsTemp("default_unit")
			bshow_on_web=rsTemp("visible")
			wh_id=rsTemp("wh_id")
			grams_default=rsTemp("grams_default")
			if bshow_on_web then bwebCheck="Checked"
			x=closeRS()
			'x=rwe(sSQL)
			%>
			<form  id="myform" name="myform" action="/admin/ajax/save.asp">
			<input type="hidden" name="t" value="food">
			<input type="hidden" name="uid" id="uid" value="<%=uid_food%>">
			<input type="hidden" name="file_name" id="file_name" value="<%=uid_food%>.jpg">
			<input type="hidden" name="folder_name" id="folder_name" value="food">
			<input type="hidden" name="idf" id="idf" value="<%=id_food%>">
			<div class="row row-centered">
				<div class="col-xs-12">
					<h3>Adding Food</h3>
				</div>
			</div>
 			<div class="row row-centered">
		        <div class="col-xs-3 indent10" >
		          Food Name
		        </div>
		        <div class="col-xs-5">
					<input type="text" size="50" maxlength="100" name="name" value="<%=name%>">        
		        </div>
		        <div class="col-xs-4">
					<a class="button icon arrowright" href="/food.asp?f=<%=id_food%>">Preview on web<a>
		        </div>


		    </div>
			<div class="row row-centered">
		        <div class="col-xs-3 indent10" >
		          Picture
		        </div>
		        <div class="col-xs-9">
				<div class="thumb-wrapper">
				<img id="update_img"  class="photo_holder_food" src="<%=Image_path%>" alt="<%=name%>" />
				<span id="updatePhoto" onclick="updatePhoto();"></span>
				<pre id="log" style="height: 300px; overflow: auto;display:none;"></pre>
				</div>         
		        </div>
		    </div>
 			<div class="row row-centered">
		        <div class="col-xs-3 indent10" >
		         	Add Vitamin
		        </div>
		        <div class="col-xs-9">
				<select name="vitamin" id ="vitamin"><option value="0">Select Vitamin</option>
						<%'enumerate foods list
						x=openRS("SELECT  v.name, v.`id_vitamin`,  v.`color` FROM `antidote`.`vitamins` v order by v.name;")
						do until rsTemp.eof%>
							<option value="<%=rsTemp("id_vitamin")%>"<%if id_type=rsTemp(0) then x=rw(" selected") %>><%=rsTemp("name")%></option>
							<%rsTemp.movenext
						loop
						x=closeRS()
						%>
					</select>    
		        % of RDI <input type="text" name="percentage" id="percentage" value="<%=percentage%>" size="2"> for serving size of   
		        <input type="text" size="6" maxsize="5" id="food_amount" name="food_amount" value="<%=grams_default%>" disabled> grams
		        <input type="button" id="add_vitamin" class="button" value="Add Vitamin" onclick="Add_Vitamin()">
		        </div>
		    </div>

 		
 			<div class="row row-centered">
		        <div class="col-xs-3 indent10" >
		          Ingredients
		        </div>
		        <div id="ingredients_list" class="col-xs-9">
		        	<%'check to see what ingredients are already added?'
						x=openRS("SELECT `id_food_vitamin`, `id_food`, v.name, v.`id_vitamin`,`percentage`,  v.`color` FROM `antidote`.`food_vitamins` fv inner join vitamins v on v.id_vitamin=fv.id_vitamin where id_food="&id_food&" order by fv.percentage desc;")
		        	if rsTemp.eof then
		        		x=rw("<b><i>Add some ingredients to your menu by using the add ingredient button above....</i></b>")
		        	else
		        		do until rsTemp.eof
		        			x=rw("<div id=""vitamin"&rsTemp("id_food_vitamin")&""">RDI % "&rsTemp("percentage")&" of "&rsTemp("name")&" <button class=""button danger icon remove""  onclick=""Delete_food_vit("&rsTemp("id_food_vitamin")&"); return false;"">Remove item</button></div>")
		        			rsTemp.movenext
		        		loop
		        	end if
		        	%>
		        </div>
		    </div>
 			<div class="row row-centered">
		        <div class="col-xs-3 indent10" style="margin-top:20px;" >
		         	Introduction
		        </div>
		        <div class="col-xs-9">
					<textarea id="intro" name="intro" class="editor"><%=Intro%></textarea>
		        </div>
		    </div>
 			<div class="row row-centered">
		        <div class="col-xs-3 indent10" style="margin-top:20px;" >
		         	Description
		        </div>
		        <div class="col-xs-9">
					<textarea id="Description" name="Description" class="editor"><%=Description%></textarea>
		        </div>
		    </div>
		    
 		    <div class="row " style="">
 		    	<div class="col-xs-4"></div>
		    	<div class="col-xs-8">
		    		 <%if session("can_authorize")=true then%>
		    		<input type="checkbox" name="show_on_web" id="show_on_web" checked="<%=bwebCheck%>"> <label for="show_on_web">Show this food to the live website.</label></br>
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
		           <button  class="button danger icon remove" onclick="Delete_Record('food','<%=uid_food%>','/foods.asp'); return false;">Delete Record</button>
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
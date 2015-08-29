<!--#include virtual="/header.htm" -->
<!-- #include virtual="/connection.asp"-->
<!-- #include virtual="/functions.asp"-->
<script type="text/javascript" src="https://www.google.com/jsapi"></script>
<div class="foods" >

<% 


'check to see if id is stated?

sSQL = "SELECT * FROM antidote.recipes where id_recipe="&request("r")&";"

x=openRS(sSQL)
%>

<%
irow=0
do until rsTemp.eof
	if irow mod 2=1 then strClass="light_blue_row" else strClass="white_row"
	irow=irow+1
	id_recipe=rsTemp("id_recipe")
	id_person=rsTemp("id_person")
	name=rsTemp("name")
	image=replace(rsTemp("image"),"med","large")
	'image=rsTemp("image")
	how_to_make=rsTemp("how_to_make")
	id_type=rsTemp("id_type")
	servings=rsTemp("servings")
	brief=rsTemp("brief")
	uid_recipe=rsTemp("uid_recipe")
	sArray=""
	%>
<script type="text/javascript">
  google.load("visualization", "1", {packages:["corechart"]});
      google.setOnLoadCallback(drawChart);
      function drawChart() {
        var data = google.visualization.arrayToDataTable([
         ['Nutrient','Link','% of RDI click bar to see more',{ role: 'style' }, { role: 'annotation' }],
   <%x=openRSA("CALL `antidote`.`Recipe_Vitamins`("&id_recipe&");")
		do until rsTempA.eof
			sArray=sArray&"['"&rstempA(1)&"','/vitamin.asp?v="&rstempA(0)&"',"&round(rstempA(2),2)&",'"&rstempA(3)&"','This meal/drink has "&round(rstempA(2),2)&" of your RDI for "&rstempA(1)&"'],"&vbcrlf
			rsTempA.movenext
		loop
		if not sArray="" then sArray=left(sArray,len(sArray)-1)
		x=closeRSA()
		x=rw(sArray)
		%>           
        ]);
       var view = new google.visualization.DataView(data);
       view.setColumns([0, 2,3]);
			 var chartAreaHeight = data.getNumberOfRows() * 20;
			// add padding to outer height to accomodate title, axis labels, etc
			var chartHeight = chartAreaHeight + 80;
       var options = {title:"% of Recommended Daily Intake for each vitamin",

            vAxis: {title: "Percentage of DRI"},
            hAxis: {title: "Nutrient"},
 						height: chartHeight,
	    			chartArea: {height: chartAreaHeight}           
            };

       var chart = new google.visualization.BarChart(document.getElementById('chart_div'));
       chart.draw(view, options);

       var selectHandler = function(e) {
          window.location = data.getValue(chart.getSelection()[0]['row'], 1 );
       }

       google.visualization.events.addListener(chart, 'select', selectHandler);
      }





</script>
<%
	'sSQL="Update food set intro='"&stripHTML(replace(rsTemp("Intro"),"'","''"))&"' where id_food='"&rsTemp("id_food")&"'"
	'x=rwb(sSQL)
	'x=openRSA(sSQL)
	'x=rwe("here.")
	%>
		<div class="row">
			<div class="col-sm-10 col-xs-10">
				<h1><a href="/recipe.asp?r=<%=id_recipe%>"><%=name%></a></h1>
			</div>
			<%if session("id_person")=id_person or session("can_authorize") then%>
			<div class="col-sm-2 col-xs-2">
				<a class="button icon edit" href="/add_recipe.asp?id=<%=id_recipe%>">Edit</a>
			</div>
			<%end if%>
		</div>
		<div class="row">
			<div class="col-sm-7 col-xs-12">
				<img src="<%=image%>" alt="name">
			</div>
			<div class="col-sm-5 col-xs-12">
				<div class="row">
			        <div class="col-xs-12 col-sm-12" >
			          <h2><%=rsTemp("brief")%></h2>
			        </div>
				</div>

				<div class="row">
			        <div class="col-xs-12 col-sm-12" >
			          <i>Ingredients</i>
			        </div>
				</div>

				<div class="row">		        
			        <div id="ingredients_list" class="col-xs-12">
			        	<ul>
			        	<%'check to see what ingredients are already added?'
			        	x=openRSA("call Recipes_By_ID ("&id_recipe&")")
			        	if rsTempA.eof then
			        		x=rw("No Ingredients")
			        	else
			        		do until rsTempA.eof
			        			x=rw("<li>"&rsTempA("qty_grams")&" grams of <a href=""/food.asp?f="&rsTempA("id_food")&""">"&rsTempA("name")&"</a></li>"&vbcrlf)
			        			rsTempA.movenext
			        		loop
			        	end if
			        	x=closeRSA()
			        	%>
			        	</ul>
			        	<div><b>Serves <%=servings%> people</b></div>
			        </div>
			    </div>
		</div>
		<div class="row">
			<div class="col-sm-6 col-xs-12">
				<h3>How to make this recipie.</h3>
				<p><%=how_to_make%></p>
			</div>
		</div>
		<div class="row">			 
				<div class="row">
			        <div class="col-sm-12 col-xs-12" >
			          <h3><a href="/vitamins.asp">Nutritional</a> breakdown (vitamins ordered by Recommended Daily Intake)</h3>
			        </div>
				</div>
			    <div id="chart_div" style="height:auto;"></div>
		</div>
	
	<%
	rsTemp.movenext
loop
x=closeRS()%>
	</div>
</div>
<div id="spacer" style="margin-top:20px;"></div>
<!--#include virtual="/footer.asp" -->
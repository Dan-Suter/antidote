<!--#include virtual="/header.htm" -->
<!-- #include virtual="/connection.asp"-->
<!-- #include virtual="/functions.asp"-->
<!--#include virtual="/security.asp" -->
<%
'this is the page where people get to comsume the recipe item that they have been viewing
'check out how many portion size options there are with this one.

'List as portion size options 
'show photo of this one + any other currently in cart...
'process person add recipes to profile.
'Show thanks you screen enjoy your meal.
'get the meal with suggested contirbution
'love_your_food.asp ASAP
'One DanaOne per meal for make at home
'Cost Plus One DanaOne per meal for in commerical service
'DanaOne per meal for in home service
'Cost is not optional
'DanaOne is optional
'add on takeaway options and danaone as checkboxes.
'Add recipes to order using ajax
'first thing we are inserting into meal time so SQL that.
if session("id_helper")="" then session("id_helper")=session("id_people")
if not request("r")="" then
	sSQL="Call Insert_people_eat ("&request("r")&" , "&session("id_people")&", "&session("id_helper")&" ,"&request("p")&" )"
	'Now show the current order from this person
	'x=rwb(sSQL)
	x=openRS(sSQL)
	x=closeRS()
	'x=rwe(sSQL)
end if
if not request("d")="" then
	sSQL="Delete from people_eat where id_people_eat="&request("d")&";"
	x=openRS(sSQL)
	x=closeRS()
end if

%>
<script>
function Delete_Meal(d)
{location.href="/loveyourfood.asp?d="+d}
</script>
<form  id="eat_me" name="eat_me" action="/givewithlove.asp">
<input type="hidden" name="uid" value="<%=uid_recipe%>">
<input type="hidden" name="idr" id="idr" value="<%=id_recipe%>">
<div class="row row-centered">
	<div class="col-xs-12">
		<h2>Hey <%=session("name")%> Whats on your plate?</h2>
	</div>
</div>
<% 

sSQL="Call get_people_eat_by_id ("&session("id_people")&")"
'x=rwb(sSQL)
x=openRS(sSQL)
itotal=0
do until rsTemp.eof
	Portion_name=rsTemp("Portion_name")
	id_person=rsTemp("id_people_eat")
	recipe_name=rsTemp("recipe_name")
	amount_currency=rsTemp("amount_currency")
	image=replace(rsTemp("image"),"med","thumb")
	itotal=itotal+amount_currency
		%>
		<div class="row row-centered">     
	     <div class="col-xs-2" style=" margin-top:10px;">
				<a class="button danger icon remove"  onclick="Delete_Meal(<%=rsTemp("id_people_eat")%>); return false;"">Remove</a>
			</div> 
	    <div class="col-xs-1">
				<img id="img_my_meal" src="<%=image%>" alt="<%=name%>" />
			</div>          
	    <div class="col-xs-8">
				<h4>1 x <%=Portion_name & " " &recipe_name %></h4>
			</div>
	    <div class="col-xs-1 text-right">
				<span class="checkout_currency"><%=FormatMoney(amount_currency)%></span>
			</div>
	  </div>
	 <%rsTemp.movenext
	Loop
	'ok so get totals and show payment links
	x=closeRS()		
	%>
		<div class="row row-centered">       
	    <div class="col-xs-8">
				<span class="checkout_currency">Enjoy your meal</span>
			</div>
	    <div class="col-xs-2">
				<span class="checkout_currency">
					<a style="min-width:70px;" class="button eatme primary" href="/contribute.asp?p=Cash">Contribute $ <%=itotal%> Cash
						</a>
					</span>
			</div>
			<div class="col-xs-2">
				<span class="checkout_currency">
					<a style="min-width:70px;" class="button danger primary" href="/contribute.aspp?p=Eftpos%>">Contribute $ <%=itotal%> Eftpos
						</a>
					</span>
			</div>
	  </div>
</form>
<div id="chart_div" style="height:auto;"></div>
<script type="text/javascript" src="https://www.google.com/jsapi"></script>


<script type="text/javascript">
  google.load("visualization", "1", {packages:["corechart"]});
      google.setOnLoadCallback(drawChart);
      function drawChart() {
        var data = google.visualization.arrayToDataTable([
         ['Nutrient','Link','% of RDI click bar to see more',{ role: 'style' }, { role: 'annotation' }],
   <%

   x=openRSA("CALL `antidote`.`Recipe_Vitamins_people_eat`("&session("id_people")&");")
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
<!--#include virtual="/footer.asp" -->




%>
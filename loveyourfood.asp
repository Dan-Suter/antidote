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
sSQL="Call Insert_people_eat ("&request("r")&" , "&session("id_people")&", "&session("id_helper")&" ,"&request("p")&" )"
'Now show the current order from this person
'x=rwb(sSQL)
'x=openRS(sSQL)
'x=closeRS()
			'x=rwe(sSQL)
%>
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
	    <div class="col-xs-2">
				<img id="update_img" src="<%=image%>" alt="<%=name%>" />
			</div>          
	    <div class="col-xs-9">
				<h3>1 x <%=Portion_name & " " &recipe_name %></h3>
			</div>
	    <div class="col-xs-1">
				<span class="checkout_currency"><%=amount_currency%></span>
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
	    <div class="col-xs-4">
				<span class="checkout_currency">
					<a style="min-width:70px;" class="button icon home" href="/loveyourfood.asp?id=<%=id_recipe%>">Contribute $ <%=itotal%> Cash
						</a>
					</span>
<span class="checkout_currency">
					<a style="min-width:70px;" class="button icon home" href="/loveyourfood.asp?id=<%=id_recipe%>">Contribute $ <%=itotal%> Eftpos
						</a>
					</span>
			</div>
	  </div>
</form>
<!--#include virtual="/footer.asp" -->




%>
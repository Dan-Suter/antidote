<!--#include virtual="/security.asp" -->
<!--#include virtual="/functions.asp" -->
<!--#include virtual="/connection.asp" -->
<%
'get URL and get Recipe food item values
if request("d")="" then
	sSQL=""
	sSQL="INSERT INTO `antidote`.`recipe_foods`(`id_Recipe`,`id_food`,`qty_grams`,`id_person`)"
	sSQL=sSQL&" VALUES	("&request("idr")&","&request("idf")&","&request("qty")&","&request("idp")&");"
	'x=rwe(sSQL)
	x=openRS(sSQL)
	x=closeRS()
	sSQL=""
	sSQL="Select id_recipe_food,qty_grams,name from recipe_foods rf inner join food f on f.id_food=rf.id_food order by id_recipe_food desc limit 1"
	x=openRS(sSQL)
	x=rw("<div id=""ingredient"&rsTemp(0)&""">"&rsTemp("qty_grams")&" grams of "&rsTemp("name")&" <button class=""button danger icon remove""  onclick=""Delete_Ingredient("&rsTemp(0)&")"">Remove item</button></div>")
	x=closeRS()
	'x=rwb("inserted "&sSQL)
else
	'must be a delete'
sSQL=""
	sSQL="Delete from  `antidote`.`recipe_foods` where id_recipe_food="&request("d")&";"
	'x=rwe(sSQL)
	x=openRS(sSQL)
	x=closeRS()
	x=rwb("Deleted "&sSQL)	
end if
%>
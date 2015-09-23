<!--#include virtual="/security.asp" -->
<!--#include virtual="/functions.asp" -->
<!--#include virtual="/connection.asp" -->
<%

if request("d")="" then
	sSQL=""
	sSQL="INSERT INTO `antidote`.`people_recipe_favourites`(`id_people`,`id_recipe`)"
	sSQL=sSQL&" Select	"&session("id_people")&","&request("r")&";"
	'x=rwb(sSQL)
	x=openRS(sSQL)
	x=closeRS()
	x=rw("favourite added")
	x=closeRS()
	'x=rwb("inserted "&sSQL)
else
	'must be a delete'
sSQL=""
	sSQL="Delete from  `antidote`.`people_recipe_favourites` where id_people="&session("id_people")&" and id_recipe="&request("r")&";"
	'x=rwe(sSQL)
	x=openRS(sSQL)
	x=closeRS()
	x=rwb("Deleted")	
end if
%>
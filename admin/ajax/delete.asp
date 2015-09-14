<!--#include virtual="/security.asp" -->
<!--#include virtual="/functions.asp" -->
<!--#include virtual="/connection.asp" -->
<%
'get URL and get Recipe food item values
if request("t")="people" then
	sSQL=""
	sSQL="Delete from `antidote`.`people`  WHERE uid_people = '"&request("d")&"';"
	x=openRS(sSQL)
	x=rw("Person Deleted redirectig to persons page")
	x=closeRS()
end if
if request("t")="recipe" then
	sSQL=""
	sSQL="Delete from `antidote`.`recipes`  WHERE uid_recipe = '"&request("d")&"';"
	x=openRS(sSQL)
	x=rw("Recipe Deleted redirectig to recipes page")
	x=closeRS()
end if
if request("t")="food" then
	sSQL=""
	sSQL="Delete from `antidote`.`food`  WHERE uid_food = '"&request("d")&"';"
	x=openRS(sSQL)
	x=rw("Food Deleted redirectig to foods page")
	x=closeRS()
end if
%>
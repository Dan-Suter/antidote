<!--#include virtual="/security.asp" -->
<!--#include virtual="/functions.asp" -->
<!--#include virtual="/connection.asp" -->
<%
'get URL and get Recipe food item values

RANDOMIZE()
Function GetRandomColor()
   Dim min, max, r, g, b
          min = 0
          max = 255

          r = Int(((max - min + 1) * Rnd) + min)
          g = Int(((max - min + 1) * Rnd) + min)
          b = Int(((max - min + 1) * Rnd) + min)
 GetRandomColor= hex(r)& hex(g)& hex(b)
'scolor="#"&GetRandomColor()
'sSQL2="update vitamins set color='"&scolor&"' where id_vitamin="&rstempA(0)&";"
'x=runUpdate(sSQL2)
End Function

if request("d")="" then
	sSQL=""
	sSQL="INSERT INTO `antidote`.`food_vitamins`(`id_food`,`id_vitamin`,`percentage`,`color`)"
	sSQL=sSQL&" Select	"&request("idf")&","&request("vitamin")&","&request("percentage")&",'#"&GetRandomColor()&"';"
	'x=rwb(sSQL)
	x=openRS(sSQL)
	x=closeRS()
	sSQL=""
	sSQL="Select id_food_vitamin,percentage,name from food_vitamins fv inner join vitamins v on v.id_vitamin=fv.id_vitamin order by id_food_vitamin desc limit 1"
	x=openRS(sSQL)
	x=rw("<div id=""vitamin"&rsTemp("id_food_vitamin")&""">RDI % "&rsTemp("percentage")&" of "&rsTemp("name")&" <button class=""button danger icon remove""  onclick=""Delete_food_vit("&rsTemp("id_food_vitamin")&"); return false;"">Remove item</button></div>")
	x=closeRS()
	'x=rwb("inserted "&sSQL)
else
	'must be a delete'
sSQL=""
	sSQL="Delete from  `antidote`.`food_vitamins` where id_food_vitamin="&request("d")&";"
	'x=rwe(sSQL)
	x=openRS(sSQL)
	x=closeRS()
	x=rwb("Deleted "&sSQL)	
end if
%>
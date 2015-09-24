<!--#include virtual="/security.asp" -->
<!--#include virtual="/functions.asp" -->
<!--#include virtual="/connection.asp" -->
<%
'get URL and get Recipe food item values

if request("d")="" then
	'sURL="/admin/ajax/add_contribution_item.asp?idc="+$("#portion_size").val()+"&c="+$("#contribuition_amount").val()+"&idr="+iR
	sSQL=""
	sSQL="call insert_recipe_contribution ("&request("idc")&",'"&request("c")&"',"&request("idr")&",0);"
	'x=rwe(sSQL)
	x=openRS(sSQL)
	x=closeRS()
	sSQL=""
	sSQL="SELECT id_recipe_contribution,  amount_currency, p.id_portion_size, p.name FROM antidote.recipe_contribution c inner join portion_sizes p	on p.id_portion_size=c.id_portion_size where id_recipe="&request("idr")&" order by id_recipe_contribution desc limit 1"
	x=openRS(sSQL)
	x=rw("<div id=""contribution"&rsTemp("id_recipe_contribution")&""">"&rsTemp("name")&" $"&rsTemp("amount_currency")&" <button class=""button danger icon remove""  onclick=""Delete_contribution("&rsTemp("id_recipe_contribution")&"); return false;"">Remove item</button></div>")
	x=closeRS()
	'x=rwb("inserted "&sSQL)
else
	'must be a delete'
	sSQL=""
	sSQL="Delete from `antidote`.`recipe_contribution` where id_recipe_contribution="&request("d")&";"
	'x=rwe(sSQL)
	x=openRS(sSQL)
	x=closeRS()
	x=rwb("Deleted "&sSQL)	
end if
%>
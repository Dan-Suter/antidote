<!--#include virtual="/header.htm" -->
<!-- #include virtual="/connection.asp"-->
<!-- #include virtual="/functions.asp"-->

<div class="welcome">
		<p><b>Welcome to Antidote.</b> Antidote is a combination of many things.  On the surface it is a small peaceful cafe in the suburb of New Brighton.
		As you look deeper at both the website and the Cafe you will find there is more than meets the eye. In terms of food 
		our menu has been created by our resident French Chef who has a passion for food and Chinese Medicene. The website has been created to be an
		open source Point of Sale system with a nutritional database engine built in.  The concept with this is to be both informative, transparent
		and to change the way in which we are relating to food. Each meal is created with thought and care with primary consideration given to the health benefits.
		For more details on this check out our <a href="/recipes.asp">recipes page</a> or to hear more about our philosophy check out the <a href="/about.asp">about page</a> 
		</p>
</div>

<% 
    'check to see if id is stated?
    '************************************************
    'update added grouping to the SQL , also added a column to determine sorting order.
    '************************************************
    bAdmin=0
    if session("can_authorize") then 
        bAdmin=1
    end if
    
    sSQL = "Call Get_recipes ("&bAdmin&");"
    x=openRS(sSQL)

	irow=0
	sFoodType=""
	bNewRow=true
	do until rsTemp.eof
		if irow mod 2=1 then
            strClass="light_blue_row" 
        else 
            strClass="white_row"
        end if    
        
		irow=irow+1
		id_recipe    = rsTemp("id_recipe")
		id_person    = rsTemp("id_person")
		name         = rsTemp("name")
        id_group_name= rsTemp("group_name")
		image        = replace(rsTemp("image"),"med","med")
		how_to_make  = rsTemp("how_to_make")
		id_type      = rsTemp("id_type")
		servings     = rsTemp("servings")
		uid_recipe   = rsTemp("uid_recipe")
        
		if sFoodType<>id_group_name then
			response.write "<div class=""group_name"">" & id_group_name	& "</div>"
            sFoodType = id_group_name
            bNewRow   = true
            irow      = 1
	    end if
        
        if bNewRow then
            response.write "<div class=""row"">"
            bNewRow=false
        end if
%>
        
        <div class="col-med-2 col-sm-4 col-xs-12">
            <div class="float_recipes">
                <a href="/recipe.asp?r=<%=id_recipe%>">
                    <img src="<%=image%>" alt="<%=name%>" class="Image_recipes">
                </a>
                <div class="Text_recipes">           
                    <a  href="/recipe.asp?r=<%=id_recipe%>"><%=name%></a>                
                </div>
            </div>
        </div>
    
<%      rsTemp.movenext
        if not rsTemp.eof then
            if iRow mod 6=0 or sFoodType<>rsTemp("group_name") Then 
                response.write "</div>"
                bNewRow=true
            end if
        else
            response.write "</div>"
        end if
    loop
    
    x=closeRS()
%>
		</div>
	</div>
</div>

<!--#include virtual="/footer.asp" -->
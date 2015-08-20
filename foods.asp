<!--#include virtual="/header.htm" -->
<!-- #include virtual="/connection.asp"-->
<!-- #include virtual="/functions.asp"-->
<div class="foods" >
	<h1>Antidote Foods</h1>
	<div class="row">
		<div class="col-md-12">
		  	<p><b>A quick google search shows there are 3.25 Billion pages with the word Food it them.</b> Anything that has needed to be said about food has surely been writen about before, and yet for the majority of us we know little about the fuel 
		  		the supplies the energy for our bodies.</p>
		  	The majority of research into food suggests that the bulk of today's dis-ease can be linked back to the quality of our diets. While we cannot overtly blame herbicides and pesticides for any particular illness or disease, there is little research
		  		 that suggests these inorganic compounds are good for you. Supposing the validity of the theory of Evolution, it would be safe to assume that our bodies have spent at least the last 99.9% of 
		  		our evolutionary path adapting to organic foods, so most of the foods listed below that we serve in our dishes is organic as well..</p>
		  		<p>What are we becoming in our haste to produce and consume more?</p>
		  		<hr>
		 </div>
	</div>
	
	
	<%if session("id_person")=id_person or session("can_authorize") then%>
		<div class="row">
				<div class="col-sm-12 col-xs-12">
					<h3>Yay you are someone who can <a class="button icon edit" href="/add_food.asp">Add A Food</a></h3>
				</div>
		</div>
	<%end if%>
	<hr>
	<% 
sSQL = "SELECT * FROM food"
x=openRS(sSQL)
irow=0
do until rsTemp.eof
if irow mod 2=1 then strClass="light_blue_row" else strClass="white_row"
irow=irow+1
image=replace(rsTemp("image_path"),"med","small")
'sSQL="Update food set intro='"&stripHTML(replace(rsTemp("Intro"),"'","''"))&"' where id_food='"&rsTemp("id_food")&"'"
'x=rwb(sSQL)
'x=openRSA(sSQL)
'x=rwe("here.")
'call Get_Image(rsTemp("image_path"),"C:\inetpub\wwwroot\antidote\images\food\"&replace(rsTemp("name"),"/","-")&".jpg")
%>
	<div class="row <%=strClass%>">
		<div class="col-md-3 col-sm-12">
			<a href="/food.asp?f=<%=rsTemp("id_food")%>"><img src="<%=image%>"></a>
		</div>
		<div class="col-md-9 col-sm-12">
			<p><a href="/food.asp?f=<%=rsTemp("id_food")%>"><%=rsTemp("name")%></a>. <%=left(rsTemp("Intro"),500)%>.... </p>
		</div>
	</div>
<%
	'x=openRSA("UPdate food set image_local=""/images/food/"&replace(rsTemp("name"),"/","-")&".jpg"" where id_food="&rsTemp("id_food"))
	'x=closeRSA()
	rsTemp.movenext
	'x=rwe("finished")

loop
'x=openRS("Call Clean_food_data")   

x=closeRS()%>
		</div>
</div>
<div id="spacer" style="margin-top:20px;"></div>
<!--#include virtual="/footer.asp" -->
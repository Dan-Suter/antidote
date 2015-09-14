<!--#include virtual="/header.htm" -->
<!-- #include virtual="/connection.asp"-->
<!-- #include virtual="/functions.asp"-->

<% 
sSQL = "SELECT * FROM food where id_food="&request("f")&";"
x=openRS(sSQL)
irow=0
if not rsTemp.eof then
	if irow mod 2=1 then strClass="light_blue_row" else strClass="white_row"
	irow=irow+1
'sSQL="Update food set intro='"&stripHTML(replace(rsTemp("Intro"),"'","''"))&"' where id_food='"&rsTemp("id_food")&"'"
'x=rwb(sSQL)
'x=openRSA(sSQL)
'x=rwe("here.")
%>
<div class="foods" >
<div class="row">
		<div class="col-xs-12">
			<ul class="imgList">
				<li class="service-list" style="width:100%">
					<img class="pad5" src="<%=rsTemp("image_path")%>" alt="icon" width="200" height="auto" />
					<h1><%=rsTemp("name")%></h1>
					<a href="/add_food.asp?id=<%=request("f")%>" class="button icon edit">Edit Food</a>
				</li>
			</ul>	
				<%
				'get the top 3 foods for each vitmain
				x=openRSA("Select fv.id_vitamin,v.name,fv.percentage/f.grams_default*100 'percentRDI',fv.color,fv.id_food from food f inner join food_vitamins fv on fv.id_food=f.id_food inner join vitamins v on v.id_vitamin=fv.id_vitamin where f.id_food="&request("f")&" order by percentage desc;")
				if not rsTempA.eof then 
					iMax=cint(rsTempA("percentRDI"))
					%>
					<table class="table table-striped">
					 <thead>
			        <tr>
			            <th>Food</th>
			            <th colspan="2">Percentage of DRI per 100 grams</th>
			        </tr>
			    </thead>
			  	<tbody>
					<%
				do until rsTempA.eof
					percent=round((cint(rsTempA("percentRDI"))/iMax)*100,0)
					%>
		        <tr>
		            <td style="width:15%"><a href="/vitamin.asp?v=<%=rsTempA("id_vitamin")%>"><%=rsTempA("name")&"</br>"%></a></td>
		            <td style="width:5%"><b><%=round(rsTempA("percentRDI"),0)%></b></td>
		            <td style="width:80%"><span class="graph" style="width:<%=percent%>%;float: left;text-indent:0px;">&nbsp;</span></td>
		        </tr>  
				<%
					rsTempA.movenext
				loop
				x=closeRSA()
				%>
							</tbody>
					</table> 
				<%end if%>
		</div>
	<div class="bodyText">	
		<div class="row">
			<div class="col-xs-12">
				<%if len(rsTemp("intro"))>0 then%>
				<%=rsTemp("intro")%>
				<%end if%>
			</div>
		</div>	
		<div class="row">
			<div class="col-xs-12">
`
				<%=rsTemp("description")%>
			</div>
		</div>	

		<div class="row">
			<div class="col-xs-12">
				<%if rsTemp("wh_id")>0 then%>
				<ul><li>Much grattidtude to George Mateljan,and the George Mateljan Foundation for <a href="http://www.whfoods.com/genpage.php?tname=foodspice&dbid=<%=rsTemp("wh_id")%>">www.whfoods.com</a></li></li></ul>
				<%end if%>
			</div>
		</div>	
	</div>
	</div>
<%
end if
x=closeRS()
%>
</div>

<div id="spacer" style="margin-top:20px;"></div>

<!--#include virtual="/footer.asp" -->
<!--#include virtual="/header.htm" -->
<!--#include virtual="/functions.asp" -->
<!--#include virtual="/connection.asp" -->
		<div class="row row-centered">
			<div class="col-sm-12">
				<h2>People of Antidote</h2>
			</div>
		</div>

		<div class="row">
				<div class="col-sm-2 col-xs-2">
					<b>Search for Person </b>
				</div>
				<div class="col-sm-10 col-xs-10">
					<input class="icon search" name="search" id="search"  placeholder="Search for a person" length="20">
				</div>
		</div>

		<div id="people_div">
			
			<%
				sSQL="Select * from people;"
				x=openRS(sSQL)
				irow=0
				do until rsTemp.eof
				if irow mod 2=1 then strClass="light_blue_row" else strClass="white_row"
				irow=irow+1
				name=rsTemp("name")
				image_path="/images/people/small/"&rsTemp("uid_people")&".jpg"
				about_me=rsTemp("about_me")
				name=rsTemp("name")
				about_me=rsTemp("about_me")
				id_people=rsTemp("id_people")
				%>
	 			<div id="htm<%=id_people%>" class="row row-centered " style="height:135px;overflow:hidden;">
	        <div class="col-xs-2" style="">
						<%if not image_path="" and not isnull(image_path) then%>
							<img src="<%=image_path%>" alt="<%=name&"'s picture"%>">
							<%
						else
							%>
							<img src="/people/images/blank-face-icon.png" alt="<%=name%>'s picture">
						<%end if%> 
						<button id="spn<%=id_people%>" class="button icon arrowdown" onclick="showMore(<%=id_people%>)">Show more.</button>          
		      </div>
	        <div class="col-xs-10">
	          <b><%=name%>.</b> <%=about_me%>
	        </div>
			  </div>
			  <hr>
			<%
				rsTemp.movenexT
			loop
			x=closeRS()
			%>
	</div>
<script type="text/javascript">

$(document).ready(
function() {
  $("#search").focus();
});
$("#search").keyup(
function() {
 $.get("/admin/ajax/search.asp?t=p&v="+$("#search").val(),function(data) 
 {$("#people_div").html(data)}
	);
});
function showMore(idv)
{
if ($("#spn"+idv).html()=="Show more."){
$("#htm"+idv).css('height', 'auto');
$("#htm"+idv).css('overflow', 'visible');
$("#spn"+idv).html("Show less.");
$("#spn"+idv).toggleClass("arrowdown arrowup");
}
else{
$("#htm"+idv).css('height', '135');
$("#htm"+idv).css('overflow', 'hidden');
$("#spn"+idv).html("Show more.");
$("#spn"+idv).toggleClass("arrowup arrowdown");}	
}
</script>		    
<!--#include virtual="/footer.asp" -->
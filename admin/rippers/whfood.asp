<!-- #include virtual="/connection.asp"-->
<!-- #include virtual="/functions.asp"-->
<%
Response.Buffer = false
'Make this page expire immediately
Response.Expires = 0
Response.Expiresabsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
             
Const Request_POST = 1
Const Request_GET = 2
response.expires=0
server.ScriptTimeout=10000

dim rsCheck
dim rsDBID
Set xobj = Server.CreateObject("MSXML2.ServerXMLHTTP")
set rsCheck=Server.CreateObject("ADoDB.recordset")
set rsDBID=Server.CreateObject("ADoDB.recordset")
iQAdded=0
sPos=""
dim sPN(10)
call OpenDB()
'get the max King_id Ripped
'x=openRS("P_Sel_KingIDMax ")
'iStart=rstEMP(0)
'x=closers()

sSearchFor=""
sSearchFor2=""
sSearchFor="<title>"

for p=125 to 300
	sPage="http://www.whfoods.com/genpage.php?tname=foodspice&dbid=foodid"
	sPage=replace(sPage,"&dbid=foodid","&dbid="&p)
	X=rwb(sPage)
	sPG=getPage(sPage)
	if instr(sPG,"<h2>The Requested Document")>0 then
		x=rwb("Nothing for : "&p)
	else
		'x=rwb("getting ... "&p)
		x=getDataFood(sPG)
	end if
next
'x=rwb(sPG)

function getDataFood(sPageData)
	'OK so the goal here is to get the data on each food item and then write that to the foods table
	dim sDescTitle
	dim sFoodName
	dim sDescription
	dim sVitaminsList
	dim sImage
	dim sTitle
	dim sIntro
	dim vitaminID(50)
	dim vitaminName(50)
	dim vitaminPercent(50)
	sFind="<title>"
	sFindEnd="</title>"
	sTitle=fnFindText(sPageData,sFind,len(sFind),sFindEnd,0,"")
	sFind="<img src="
	sPageData=startFromText(sPageData,sFind,1)
	sPageData=startFromText(sPageData,sFind,1)
	sImage="http://www.whfoods.com/"
	sPath=fnFindText(sPageData,sFind,len(sFind)+1,""" ",0,"")
	sImage=sImage&sPath
	sFind="<div style=""background:rgba(255,208,71,.7);margin:3px 0;padding:8px; font-size:3em;font-weight:bold"">"
	sFinish="</div>"
	sPageData=startFromText(sPageData,sFind,0)
	sFoodName=fnFindText(sPageData,sFind,len(sFind),sFinish,0,"")
	sPageData=startFromText(sPageData,sFoodName,1)
	'food nutrution data
	'intro.....
	sFind="<p>"
	sFinish="<br style=""clear:both"">"	
	sPageData=startFromText(sPageData,sFind,0)
	sIntro=fnFindText(sPageData,sFind,len(sFind),sFinish,0,"")
	sPageData=startFromText(sPageData,sFoodName,1)
	sPageData=startFromText(sPageData,sFinish,1)
	'get Table of nutritional info table.
	sPageData=startFromText(sPageData,"Nutrient</span>",1)
	sPageData=startFromText(sPageData,sFinish,1)
	'
	sFind="genpage.php?tname=nutrient&amp;dbid="
	sFinish=""">"	
	'x=rwe(instr(sPg,sFind))
	k=0
	while instr(sPageData,sFind)<300 and instr(sPageData,sFind)>0 and k<100
		vitaminID(k)=fnFindText(sPageData,sFind,len(sFind),sFinish,0,"")
		sPageData=startFromText(sPageData,sFind&vitaminID(k),1)
		vitaminName(k)=fnFindText(sPageData,">",1,"</a>",0,"")
		sPageData=startFromText(sPageData,"</a>",1)
		vitaminPercent(k)=fnFindText(sPageData,""">",2,"</span>",-1,"")		
		'ok so time to write these into the vitamin table and food_vitamins table
		if vitaminID(k)="" then
			k=99
		end if
		k=k+1
		sFind="genpage.php?tname=nutrient&amp;dbid="
		'x=rwb(instr(sPageData,sFind))
	wend

	sDecription=left(sPageData,instr(sPageData,"<div class=""slot-9"">")-15)
	'sDecription=left(sDecription,len(sDecription)-12)		
	'insert data into food table but check first food type does not exist
	x=openRS("SELECT id_food FROM antidote.food where name='"&dbl_apos(sTitle)&"';")
	if rsTemp.eof then
		x=closeRS()
		sSQL="INSERT INTO `antidote`.`food` (`wh_id`,`name`,`Intro`,`Image_path`,`Image_local`,`Description`)"
		sSQL=sSQL&"VALUES ("&p&",'"&dbl_apos(sTitle)&"','"&dbl_apos(sIntro)&"','"&dbl_apos(sImage)&"','"&dbl_apos(sTitle)&"','"&dbl_apos(sDecription)&"');"
		x=openRS(sSQL)
		x=closeRS()
		x=openRS("SELECT id_food FROM antidote.food where name='"&dbl_apos(sTitle)&"';")
		iFoodID=rsTemp(0)
	else
		iFoodID=rsTemp(0)
	end if
	'iFoodID is the new or old ID
	'now check for vitamins
	'if no vitamin exist create vitamin else get the vitaminID and insert to the food vitamins table...
	'check data is as expected only for Development
	'x=rwb("time to get some vitamins" & ubound(vitaminID))
	for i=0 to ubound(vitaminID)-1
		if not vitaminID(i)="" then
			x=rwb("Checking "&vitaminID(i))
			x=openRS("SELECT id_vitamin FROM antidote.vitamins where name='"&dbl_apos(vitaminName(i))&"';")
			if rsTemp.eof then
				'vitamin 
				sPageVit=getPage("http://www.whfoods.com/genpage.php?tname=nutrient&dbid="&dbl_apos(vitaminID(i))&"")
				x=closeRS()

				sDescription=fnFindText(sPageVit,"<table border=0>",0,"<div class=""slot-9"">",-15,"")
				'x=rwe(sDescription) 
				'x=rwe(instr(sDescription,"<h2>"))
				sBasicDescription=startFromText(sDescription,"</table>",1)
				sBasicDescription=fnFindText(sBasicDescription,"<h2",0,"<h2",0,"")
				'x=rwe(sBasicDescription)
				sSQL="INSERT INTO `antidote`.`vitamins` (`name`,`whf_id`,`Overview`,`Full_Description`)"
				sSQL=sSQL&"VALUES ('"&dbl_apos(vitaminName(i))&"','"&vitaminID(i)&"','"&dbl_apos(sBasicDescription)&"','"&dbl_apos(sDescription)&"');"
				'x=rwb(sSQL)
				x=openRS(sSQL)
				x=closeRS()
				x=openRS("SELECT id_vitamin FROM antidote.vitamins where name='"&dbl_apos(vitaminName(i))&"';")
				iVitID=rsTemp(0)
			else
				iVitID=rsTemp(0)
			end if
			x=closeRS()
			'awesome now we have the vit_id and food_id so we can write to food_vitamins
			x=openRS("SELECT id_food_vitamin FROM antidote.food_vitaimns where id_food="&iFoodID&" and id_vitamin="&iVitID&";")
			if rsTemp.eof then
				x=closeRS()
				sSQL="INSERT INTO `antidote`.`food_vitaimns` (`id_food`,`id_vitamin`,`percentage`)"
				sSQL=sSQL&"VALUES ('"&iFoodID&"','"&iVitID&"','"&vitaminPercent(i)&"');"
				'x=rwb(sSQL)
				x=openRS(sSQL)
				x=closeRS()
				x=openRS("SELECT id_food_vitamin FROM antidote.food_vitaimns where id_food="&iFoodID&" and id_vitamin="&iVitID&";")
				iFoodvit=rsTemp(0)
			else
				iFoodvit=rsTemp(0)
			end if	
			'x=rwb("Nice one inserted new combo for For this food :"&iFoodID&", and Vit:"&iVitID&", Percentage RDI:"&vitaminPercent(i))
		else
			i=ubound(vitaminID)
		end if
	next
	x=rwb("Yes rock and roll baby data is inserted for "&sFoodName&"! <img src="""&sImage&""">")
end function

function dbl_apos(txt)
dbl_apos=replace(txt,"'","''")
end function


%>
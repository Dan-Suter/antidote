<%'quick bit of code to update images to use UID value'
x=openRS("SELECT uid_food,image_path,name FROM antidote.food where id_food>0;")
do until rsTemp.eof
	'rename files'
	imagepath=sFilePath&replace(rsTemp(1),rsTemp(0),rsTemp(2))
	replacepath=replace(imagepath,rsTemp(2),rsTemp(0))
	x=ReNameFile(imagepath,replacepath)
	y=rwb(x)
	imagepath=replace(imagepath,"/food/small/","/food/original/")
	replacepath=replace(replacepath,"/food/small/","/food/original/")
	x=ReNameFile(imagepath,replacepath)
	y=rwb(x)
	imagepath=replace(imagepath,"/food/original/","/food/xlarge/")
	replacepath=replace(replacepath,"/food/original/","/food/xlarge/")
	x=ReNameFile(imagepath,replacepath)
	y=rwb(x)
	imagepath=replace(imagepath,"/food/xlarge/","/food/large/")
	replacepath=replace(replacepath,"/food/xlarge/","/food/large/")
	x=ReNameFile(imagepath,replacepath)
	y=rwb(x)
	imagepath=replace(imagepath,"/food/large/","/food/med/")
	replacepath=replace(replacepath,"/food/large/","/food/med/")
	x=ReNameFile(imagepath,replacepath)
	y=rwb(x)
	imagepath=replace(imagepath,"/food/med/","/food/small/")
	replacepath=replace(replacepath,"/food/med/","/food/small/")
	x=ReNameFile(imagepath,replacepath)
	y=rwb(x)
	imagepath=replace(imagepath,"/food/small/","/food/thumb/")
	replacepath=replace(replacepath,"/food/small/","/food/thumb/")
	x=ReNameFile(imagepath,replacepath)
	y=rwb(x)
	imagepath=replace(imagepath,"/food/thumb/","/food/xsthumb/")
	replacepath=replace(replacepath,"/food/thumb/","/food/xsthumb/")
	x=ReNameFile(imagepath,replacepath)
	y=rwb(x)
	sSQL="update `antidote`.`food` set image_path='/images/food/small/"&rsTemp(0)&".jpg' where uid_food='"+rsTemp(0)+"';"
	x=openRSA(sSQL)

	'y=rwe(sSQL)
	x=closeRSA()
	rsTemp.movenext
loop


%>
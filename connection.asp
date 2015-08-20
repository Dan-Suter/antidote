
<%

dim sMailerDB
dim conn
dim ConnA
dim connM
dim sMDB
dim RsTemp
dim RsTempA
'dim ipaddress
ipaddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
if ipaddress = "" then
    ipaddress = Request.ServerVariables("REMOTE_ADDR") 
end if
strDatabase = "Driver={MySQL ODBC 3.51 Driver}; Server=localhost;Port=3306;Database=antidote;User=antidote2;Password=antidote;Option=3;"
sMDBA = "Driver={MySQL ODBC 3.51 Driver}; Server=localhost;Port=3306;Database=antidote;User=root;Password=Strength2010;Option=3;"
sMDB=strDatabase


Sub OpenDB()
set Conn=server.createobject("ADODB.connection")
set RsTempC=server.createobject("ADODB.recordset")
Conn.Connectionstring=sMDB
Conn.open
End Sub

Sub CloseDB()
on error resume next
RsTempC.close
set RsTempC=nothing
Conn.close
Set Conn=Nothing
on error goto 0
End Sub


Sub CloseDBM()
on error resume next
RsTemp.close
set RsTemp=nothing
ConnM.close
Set ConnM=Nothing
on error goto 0
End Sub

Sub OpenDBM()
set ConnM=server.createobject("ADODB.connection")
ConnM.Connectionstring=sMailerDB
ConnM.open
End Sub


Sub OpenDBA()
set ConnA=server.createobject("ADODB.connection")
ConnA.Connectionstring=sMDBS
ConnA.open
End Sub

Sub CloseDBA()
on error resume next
ConnA.close
Set ConnA=Nothing
on error goto 0
End Sub
%> 








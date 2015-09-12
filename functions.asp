<%

Set xobj = Server.CreateObject("MSXML2.ServerXMLHTTP")
'******************************************
'Pay DPS Payment Upgrade 14/04/2015
'By Dan. replace pages myinvoices.asp,
'******************************************
function Pay_DPS(sXmlAction)

Dim objXMLhttp 
Set objXMLhttp = server.Createobject("MSXML2.ServerXMLHTTP.6.0") 
objXMLhttp.Open "POST", "https://sec.paymentexpress.com/pxpay/pxaccess.aspx" ,False
objXMLhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objXMLhttp.send sXmlAction
if err.number<>0 then
		on error goto 0
			x=rda("/payments/dpsServerFail.asp?err="&err.number)
		end if
on error goto 0
Dim objXML, URI
'Set oXML = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
Set objXML = CreateObject("Msxml2.DOMDocument.6.0")
ObjXML.async=true
xml=objXMLhttp.responseText
objXML.load(xml)
'function FnFindText(sSearchText,sFindText,sFindOffset,sEndText,sEndOffset,sRegExp)
URI=fnFindText(xml,"<URI>",5,"</URI>",0,"")
'x=rwe(URI)
'URI = objXML.selectSingleNode("//URI").text
Response.Redirect (replace(URI,"&amp;","&"))
Set objXMLhttp = nothing
end function


Function stripText(HTMLstring,sRegExpr)
Set RegularExpressionObject = New RegExp
With RegularExpressionObject
'.Pattern = "<[^>]+>"
.Pattern=sRegExpr
.IgnoreCase = True
.Global = True
End With
stripText = RegularExpressionObject.Replace(HTMLstring, "")
Set RegularExpressionObject = nothing
End Function

Function fncLastDay(intMonth,intYear)

	Dim intDay

	Select Case intMonth
		Case 1, 3, 5, 7, 8, 10, 12
			intDay = 31
		Case 4, 6, 9, 11
			intDay = 30
		Case 2
			If intYear mod 4 = 0 Then
				If intYear mod 100 = 0 AND intYear mod 400 <> 0 Then
					intDay = 28
				Else
					intDay = 29
				End If
			Else
				intDay = 28
			End If
	End Select

	fncLastDay = intDay

End Function

function getTMOrderFromInvoiceID(sInvoiceID)
	'get last order for customer paid by Dps that order paid_confirmed is null
	x=openRS("P_SEL_Order_From_TM_Invoice '"&sInvoiceID&"'")
	'x=rwe("P_SEL_Order_From_CUSPAYID '"&sCusPayID&"',"&idb)
	if not rstemp.eof then
		iOr=rsTemp(0)
	else
		iOr=0
	end if
	x=closeRS()
	getTMOrderFromInvoiceID=iOr
end function

function getTMPurchaseIDFromTMID(iTM)
	'get last order for customer paid by Dps that order paid_confirmed is null
	x=openRS("P_SEL_TM_PID_From_TM_ID '"&iTM&"'")
	'x=rwe("P_SEL_Order_From_CUSPAYID '"&sCusPayID&"',"&idb)
	if not rstemp.eof then
		iPID=rsTemp(0)
	else
		iPID=0
	end if
	x=closeRS()
	getTMPurchaseIDFromTMID=iPID
end function

Function IsValidEmail(sEmail)
  IsValidEmail = false
  Dim regEx, retVal
  Set regEx = New RegExp
  ' Create regular expression:
  regEx.Pattern ="^[\w-\.]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,3}$"
  ' Set pattern:
  regEx.IgnoreCase = true
  ' Set case sensitivity.
  retVal = regEx.Test(sEmail)
  ' Execute the search test.
  If not retVal Then
    exit function
  End If
  IsValidEmail = true
End Function

function getOrderFromCustomer(sCusPayID,iTrans)
'get last order for customer paid by Dps that order paid_confirmed is null
x=openRS("P_SEL_Order_From_CUSPAYID '"&sCusPayID&"',"&iTrans&","&iDB)
'x=rwe("P_SEL_Order_From_CUSPAYID '"&sCusPayID&"',"&idb)
if not rstemp.eof then
	iOr=rsTemp(0)
else
	iOr=0
end if
x=closeRS()
getOrderFromCustomer=iOr
end function

sub payByDPS(iCus,iOrder,iamount,iP,tType,iDB)
iamount=replace(formatnumber(iamount,2),",","")
sSQL="P_Ins_DPS_References '"&iP&"','"&tType&"','"&iamount&"','"&iCus&"','"&or_number&"',"&iDB
'x=rwe(sSQL)
x=openRS(sSQL)
if not rsTemp.eof then  
	x=rwe(rsTemp(0))
	sXmlAction = sXmlAction & "<GenerateRequest>"
	sXmlAction = sXmlAction & "<PxPayUserId>Laptopbattery</PxPayUserId>"
	sXmlAction = sXmlAction & "<PxPayKey>34430f1785c181141b464bbda71ab01134742c97be718c866d3f8efdd61d780b</PxPayKey>"
	sXmlAction = sXmlAction & "<TxnType>Purchase</TxnType>"
	sXmlAction = sXmlAction & "<CurrencyInput>NZD</CurrencyInput>"
	sXmlAction = sXmlAction & "<AmountInput>"&iamount&"</AmountInput>"
	sXmlAction = sXmlAction & "<MerchantReference>"&rsTemp(0)&"</MerchantReference>"
	sXmlAction = sXmlAction & "<EmailAddress>sales@laptopbattery.co.nz</EmailAddress>"
	sXmlAction = sXmlAction & "<UrlSuccess>"&sSiteURL&"/dpsresponse.asp</UrlSuccess>"
	sXmlAction = sXmlAction & "<EnableAddBillCard>1</EnableAddBillCard>"
	sXmlAction = sXmlAction & "<UrlFail>"&sSiteURL&"/dpsresponse.asp</UrlFail>"
	sXmlAction = sXmlAction & "</GenerateRequest>"	
	'response.write sXmlAction
	'response.end
	Dim objXMLhttp 
	Set objXMLhttp = server.Createobject("MSXML2.XMLHTTP") 
	objXMLhttp.Open "POST", "https://www.paymentexpress.com/pxpay/pxaccess.aspx" ,False
	objXMLhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXMLhttp.send sXmlAction
	Dim oXML, URI
	Set oXML = Server.CreateObject("MSXML2.DomDocument")
	oXML.loadXML(objXMLhttp.responseText)
	URI = oXML.selectSingleNode("//URI").text
	Response.Redirect URI
	Set objXMLhttp = nothing
end if
end sub
	


function ifs()
ifs=false
if not session("staff_id")="" then
	ifs=true
end if
end function


function getCustomerFromOrder(iOrderNo,iDBO)
x=openRS("P_SEL_OrderTotal "&iOrderNo&","&iDBO)
if not rsTemp.eof then
	getCustomerFromOrder=rsTemp(1)
	iOrderNo=rsTemp(2)
else
	x=rwbe("<span class=""attention"">No orders found for the order number: "&iOrderNo&"</span>")
end if
x=closeRS()
end function


function GetTransData(iTransID)
iAmount=request("LBT_Amount")
if iAmount="" then iAmount=0
set rsObject=server.createobject("ADODB.recordset")
'x=rwe("P_Sel_LBTRans "&iTransID&","&idb)
rsObject.open "P_Sel_LBTRans "&iTransID&","&idb,sMDB
'response.write "P_Sel_LBTRans "&iTransID&","&idb&"</br>"
sHTML="<div id=""transR""><div class=""trHead"">Results for Trans ID: "&iTransID&"</div>"
sHTML=sHTML& "<table class=""tblTrans"" ><tr class=""qFieldNames""><td></td><td>Date:</td><td>Category</td><td>Customer</td><td>Reference</td><td align=right>Amount</td><td>Note</td></tr>"
if not rsObject.eof then
	bDisplayTrans=true
end if
i=2

do until rsObject.eof
	irow=irow+1
	if irow mod 2=1 then strClass="light_blue_row" else strClass="white_row"
	if isnull(rsObject("LBT_Amount")) then iAmount=0 else iAmount=rsObject("LBT_Amount")
		sRefLink=""
		if rsObject("purchaseOrder")>0 THEN
			sRefLink="<a href=""/admin/OnOrderNew.asp?k="&rsObject("purchaseOrder")&""">PO"&rsObject("purchaseOrder")&"</a>"
		end if
		if rsObject("LBT_OrID")>0 then
			sRefLink="<a href=""/order_man.asp?Or_Number="&rsObject("LBT_OrID")&""">Order "&rsObject("LBT_OrID")&"</a>"
		end if			
	sHTML=sHTML& "<tr class="""&strClass&"""><td onmouseover=""this.style.cursor='pointer'"" onclick=""DelTrans("&rsObject("LBT_ID")&")"">"
	sHTML=sHTML& "<font color=""red"">x</font></a></td><td>"&CvbShortdate(rsObject("LBT_Filed_date"))&"</td>"
	sHTML=sHTML& "<td nowrap>"&rsObject("CT_Text")&"</td>"
	sHTML=sHTML& "<td align=""left""><a href=""/order_man.asp?cu_name="&rsObject("LBT_CuID")&""">"&rsObject("Cus_First_Name")&" " &rsObject("Cus_Last_Name")&"</td>"
	sHTML=sHTML& "<td align=""right"">"&sRefLink&"</td><td align=""right"">"&formatNumber(iAmount,2)&"</td><td><input class=""button"" type=""text"" name=""note_"&rsObject("LBT_ID")&""" value="""&rsObject("LBT_Note")&""" onkeydown=""EditNote("&rsObject("LBT_ID")&",event,this.value);cancelBubble(event)""></td></tr>"
	rsObject.movenext
loop

rsObject.close
rsObject.open "P_Sel_CompareLBTandBTTotals "&iTransID&","&idb,sMDB
'response.write "P_Sel_CompareLBTandBTTotals "&iTransID&","&idb &"</br>"
'if round(rsObject(0),2)=round(rsObject(1),2) then
'	Call Reconcile(iTransID,0,"<font color=green><b>Filed to "&sCatText&" Trans Total :$"&rsObject(0)&" Bank $"&rsObject(1)&" and RECONCILED</b></font>",1)
'else
'	Call Reconcile(iTransID,0,"<font color=red><b>Filed to "&sCatText&" Trans Total :$"&rsObject(0)&" Bank $"&rsObject(1)&" and NOT RECONCILED</b></font>",0)
'end if

if isnull(rsObject(0)) then iAmount=0 else iAmount=rsObject(0)

sHTML=sHTML& "<tr><td colspan=6 align=left>--------------------------------------------------------------------------------</td></tr>"
sHTML=sHTML& "<tr><td colspan=5 align=right>Transactions Total:</td><td align=""right"">$"&formatNumber(iAmount,2)&"</u></b></td></tr>"
sHTML=sHTML& "<tr><td colspan=5 align=right>Related Bank Amount:</td><td align=""right"">$"&formatNumber(rsObject(1),2)&"</u></b></td></tr>"
sHTML=sHTML& "<tr><td colspan=5 align=right>Balance to Reconcile:</td><td align=""right""><b><u>$"&formatNumber(rsObject(1)-iAmount,2)&"</u></b></td></tr>"
sHTML=sHTML& "</table>"
rsObject.close

set rsObject=nothing

response.write sHTML
response.end
end function

function fnTurnSQLtoSessionVars(sSQL)
x=rsOpen(sSQL)
for each fld in rsTemp.fields
	session(fld.name)=rsTemp(fld.name)
next
end function

function ValidateFieldsUser(rsDD,sKey)
'check data if it passes validation then allow save.
'x=rwe("")
dim iErrors
bFieldPass=true
with rsDD
do until.eof
	if .fields("IS_Shown") then
		iFldID=.fields("FL_ID")
		'Check field by Field
		sFieldName=.fields("field_name")
		sValue=request(replace(.fields("field_friendly_name")," ","_"))
		sValidation=.fields("field_type_des")
		sValText=.fields("Validation_Text")
		iMaxLength=.fields("field_length")
		bReq=.fields("Required")
		sPass=false
	 	'x=rwb(sFieldName&":"&sValue)
	 	if rsDD("key_field")=true then iKeyFieldID=rsDD("fl_ID")
		Select case sValidation
			case "bit"
				sPass=true
			case "int"
		if rsDD("field_name")="msg_to" then
			'response.write sPass&sValue
			'response.end
		end if
				if isnumeric(sValue)=true then
					if sValue>-2147483648 and sValue<2147483647 then
						sPass=true					
					else
						sErrCode=sErrCode & "<e"&iFldID&">" & sValText & "</e"&iFldID&">"
						iErrors=iErrors+1
					end if
				else
					if bReq=false and sValue="" then
						sPass=true			
					else
						sErrCode=sErrCode & "<e"&iFldID&">" & sValText & " and is Required</e"&iFldID&">"
						iErrors=iErrors+1
					end if
				end if
			case "smallint"
				if isnumeric(sValue)=true then
					if sValue>-32768 and sValue<32768 then
						sPass=true					
					else
						sErrCode=sErrCode & "<e"&iFldID&">" & sValText & "</e"&iFldID&">"
						iErrors=iErrors+1
					end if
				else
					if bReq=false and sValue="" then
						sPass=true			
					else
						sErrCode=sErrCode & "<e"&iFldID&">" & sValText & " and is Required</e"&iFldID&">"
						iErrors=iErrors+1
					end if
				end if
			case "smallint"
				if isnumeric(sValue)=true then
					if sValue>-1 and sValue<256 then
						sPass=true					
					else
						sErrCode=sErrCode & "<e"&iFldID&">" & sValText & "</e"&iFldID&">"
						iErrors=iErrors+1
					end if
				else
					if bReq=false and sValue="" then
						sPass=true			
					else
						sErrCode=sErrCode & "<e"&iFldID&">" & sValText & " and is Required</e"&iFldID&">"
						iErrors=iErrors+1
					end if
				end if	
			case "date","datetime","smalldatetime"
				if isdate(sValue)=true then
					sPass=true				
				else
					if bReq=false and sValue="" then
						sPass=true
					else
						sErrCode=sErrCode & "<e"&iFldID&">This field must have a date value.</e"&iFldID&">"
						iErrors=iErrors+1
					end if
				end if
			case else 
				if len(sValue)>iMaxLength then
					sPass=false
					sErrCode=sErrCode & "<e"&iFldID&">You have entered to much text.  Maximum allowed is "&iMaxLength&". Please reduce by "&abs(iMaxLength-len(sValue))&" characters.</e"&iFldID&">"
				else
					if bReq=true and (sValue="" or sValue="</br>") then
						if session("staff_id")=1 then
							sErrCode=sErrCode & "<e"&iFldID&"><a href=""/admin/item.asp?t=19&id="&iFldID&""">This field must have a value.</a></e"&iFldID&">"
						else
							sErrCode=sErrCode & "<e"&iFldID&">This field must have a value.</e"&iFldID&">"
						end if
						iErrors=iErrors+1
					else
						sPass=true			
					end if		
				end if
		end select
	
		if sPass=false then bFieldPass=false
		if sPass=false then 
				'response.write "No Pass for :" & sFieldName & "=" & sValue &"</br>"
				'response.end
		end if
	end if
	.movenext
loop
rsDD.movefirst 
'response.write sPassVal
'response.end
end with
ValidateFieldsUser=sErrCode
'x=rwe(sErrCode)
if ValidateFieldsUser<>"" then 
	ValidateFieldsUser="<validate><failed><table class=""tResultsFail""><tr><td><img src=""/styles/images/error.gif""></td><td><span class=""error"" />Please Check the form "&iErrors&" errors were found.</span></td></tr></table></failed>"&ValidateFieldsUser&"</validate>"
else 
	ValidateFieldsUser=""
end if
end function

function getLastID(iTbl)
'get last inserted ID for DB_owner
x=openRS("P_Sel_LastIDByTable "&iTbl&","&idb)
if not rsTemp.eof then
	getLastID=rsTemp(0)
else
	getLastID=0
end if
end function

function saveCusDataFldName(iField,iTbl,iGroup)
sDBO=""
sKey=stripinvalid(request("key"))
'Allows edit of cusData via field,table, optional group limiter use 0 for all groups
set rstDD=server.createobject("adodb.recordset")
sSQl="P_Sel_FieldLookup_CusEdit "&iField&","&iTbl&","&iGroup
'x=rwb(sSQL)
rstDD.open sSQL,sMDB
if rstDD.eof then
	response.write "EOF"
	response.end
end if
'now with rs getData
set rs=server.createobject("adodb.recordset")
sValidated=ValidateFieldsUser(rstDD,sKey)
if rstDD("has_db_owner") then	sDBO=idb
if len(sValidated)>0 then 
	'must be an error
	response.write (sValidated)
	response.end 
end if
if request("Action")="1" then
	'insert action
	sSQL="P_T_INS_"&rstDD("table_name")&"_By_CUS_SYS "
	sSQL=sSQL&session("cus_sys")&","
else
	'update action
	sSQL="P_T_UP_"&rstDD("table_name")&"_By_CUS_SYS "
	sSQL=sSQL&session("cus_sys")&","&sKey&","
end if
do until rstDD.eof
	if rstDD("is_shown") and not rstDD("key_field") then
		sVal=request(replace(rstDD("field_friendly_name")," ","_"))
		if rstdd("field_type_des")="bit" then
			if sVal="true" then 
				sVal=1
			else
				sVal=0
			end if
		end if
		sSQL=sSQL&"'"&sVal&"',"
		'x=rwe("i"&sKey&"f"&rstDD("fl_id"))
	end if
	rstDD.movenext
loop
'get all fields
sSQL=sSQL&sDBO
if right(sSQL,1)="," then sSQL=left(sSQL,len(sSQL)-1)
'x=rwb(sSQL)
on error resume next
'on error goto 0
x=openRS(sSQL)
for each fld in rsTemp.fields
	x=rwb(fld.value)
next
if err.number=0 then
	if iTbl=13 then
		x=rwb("<div class=""myInfoWarning""><div class=""errorIcon"">&nbsp;</div>Please note you will need to logout and log back in for some of the udpated changes to affect your account.</div>")
	end if
	x=rwb(GetSavedLinks()&"</br>"&WhereToLinks())
else
	'x=rwe(sSQL)
	x=rwe(GetErrorOccured(sSQL))
end if
'on error goto 0
end function

function fixUid(sUID)
fixUid=replace(replace(sUID,"{",""),"}","")
end function


function GenerateEmailBody(eName)

	x=openRSA("P_Sel_System_Templates_Body '"&eName&"',"&idb)
	'x=rwb(rsTempA.source)
	iEid=rsTempA("id")
	sBody=rsTempA("body")
	sSub=rsTempA("subject")
	x=closeRSA()
 
GenerateEmailBody=sBody
end function

Function SendSystemMail(eName,iOr,iCu,bForceSend)
'exit function
iOr=cdbl(iOr)
iCu=cdbl(iCu)
set rsEQ=server.createobject("ADODB.recordset")
x=openRSA("P_SEL_Email_Sys_Content '"&eName&"',"&idb)
'x=rwe(rsTempA.source)
iEid=rsTempA("id")
sBody=rsTempA("body")
'x=rwe(sBody)
sSub=rsTempA("subject")
x=closeRSA()
sCusTo=EmailsList(eName,iOr,iCu)
sFrom=EmailGetSysEml(eName,idb)
'x=rwe(EmailGetAttachments(eName,iOr,iCu,iDB))
iAttach=EmailGetAttachments(eName,iOr,iCu,iDB)
'x=rwe(sSub)
sSub=EmailConvertBody(sSub,iOr,iCu)
x=OpenRSA("P_Sel_EmailHeader_and_Footers "&idb)
sHeader=""
sFooter=""
if not rsTempA.eof then
sHeader=rsTempA(0)
sFooter=rsTempA(1)
end if
x=closeRS()
'if not left(sBody,4)="<?xml" and instr(sBody,"<html")=0 then
'	'sHeader=replace(sHeader,"[[DB_SiteTradingName]]",sSiteTradingName)
'end if
sBody=sHeader+sBody+sFooter
sBody=EmailConvertBody(sBody,iOr,iCu)
sCusEmails=split(sCusTo,",")
iEmails=ubound(sCusEmails)
'bFormat=session("CO_Invoice_In_HTML")
bFormat=true
if len(iAttach)=0 then iAttach=0
	istaff=session("staff_id")
if session("staff_id")="" then istaff=0
for i=0 to iEmails
	sSQL="P_Ins_EmailQueSys '"&sFrom&"','"&sCusEmails(i)&"','"&replace(sSub,"'","''")&"','"&replace(sBody,"'","''")&"','"&bFormat&"',NULL,"&iStaff&",'sys."&eName&"',"&iAttach&","&iEid&","&iDB
	'x=rwe(sSQL)
	'x=ifa(i&":"&sSQL)
	rsEQ.open sSQL,sMDBA
	sResponse=sResponse&rsEQ(0)
	rsEQ.close
	if instr(sResponse,"email successfully")>0 then
		iEmails=iEmails+1
	else
		sErr=sErr&sResponse&" </br>"
	end if
	sSQL=""
	x=rw(sResponse)
	'x=ifa(sResponse)
next			
set rsEQ=nothing
'x=rwe("finsh send mail check your inbox.")
end Function


function EmailsList(eName,iOr,iCu)
sE=""
'x=rwb(iOr)
if not iOr=0 then
	'get customer details from order
	x=openRS("P_Sel_Order_CusEmailByOrder "&iOr&","&iDB)
	if not rsTemp.eof then 
		sCusName=rsTemp("cus_first_name")& " "& rsTemp("cus_last_name")
		sCusEmail=rsTemp("cus_email")
		sCusAccName=rsTemp("cus_accounts_name")
		sCusAccEml=rsTemp("cus_Accounts_email")
	end if
	x=CloseRS()
end if
if len(sCusEmail)=0 and iCu>0 then
		x=OpenRS("P_Sel_CustomerByID "&iCu&","&iDB)
		if not rsTemp.eof then
			sCusName=rsTemp("cus_first_name")& " "& rsTemp("cus_last_name")
			sCusEmail=rsTemp("cus_email")
			sCusAccName=rsTemp("cus_accounts_name")
			sCusAccEml=rsTemp("cus_Accounts_email")
		end if
		x=CloseRS()
end if
sCusName=replace(sCusName,",","_")
sCusAccEml=replace(sCusAccEml,",","_")
if not sCusEmail="" then
	if eName="order_confirm" or eName="payment_confirm" then
		sE=sE&""""&sCusName&""" <"&sCusEmail&">"
		if len(sCusAccEml)>0 then sE=sE&","""&sCusAccName&""" <"&sCusAccEml&">"
	end if
	if eName="tracking_confirm" then
		sE=sE&""""&sCusName&""" <"&sCusEmail&">"
	end if
end if
if sE="" then sE=""""&sCusName&""" <"&sCusEmail&">"
EmailsList=sE 
'x=rwe (EmailsList)
end Function

function createSysFolders(iDB2)
set fso = CreateObject("Scripting.FileSystemObject")
'x=rwe("\images\DB" &iDB2 &"\Stock\")
f=fso.createfolder(sSitePath&"\images\DB" &iDB2 &"\Stock\")
f=fso.createfolder(sSitePath&"\images\DB" &iDB2 &"\Stock\TM_Photos")
f=fso.createfolder(sSitePath&"\images\DB" &iDB2 &"\Stock\Original\")
f=fso.createfolder(sSitePath&"\images\DB" &iDB2 &"\Stock\x-large\")
f=fso.createfolder(sSitePath&"\images\DB" &iDB2 &"\Stock\large\")
f=fso.createfolder(sSitePath&"\images\DB" &iDB2 &"\Stock\normal\")
f=fso.createfolder(sSitePath&"\images\DB" &iDB2 &"\Stock\small\")
f=fso.createfolder(sSitePath&"\images\DB" &iDB2 &"\Stock\x-small\")
f=fso.createfolder(sSitePath&"\images\DB" &iDB2 &"\Stock\xx-small\")
set fso=nothing
end function

function makepath(sPath)
dim fs,f
set fs=Server.CreateObject("Scripting.FileSystemObject")
if not fs.folderExists(sPath) then
'x=rwe(sPath)
set f=fs.CreateFolder(sPath)
end if
set f=nothing
set fs=nothing
end function

sub openFSO() 
    set fso = CreateObject("Scripting.FileSystemObject") 
end sub 

sub closeFSO() 
    set fso = nothing 
end sub 


Function EmailGetAttachments(sEmlType,iOrd,iCus,iDB)
iAtachments=0
Select case sEmlType
	case "order_confirm"
		'pdf invoice and serial list coming up
		'x=rwe("P_sel_OrderUIDFromOrNo "&iOrd&","&iDB)
		x=CreatePDFfromInvoice(iOrd)
		EmailGetAttachments=x
	case "payment_confirm"
		'nothing for you
		exit function
	case "tracking_confirm"
		'PDF of delivery confirmation
		exit function
	case else
		exit function
end Select
end Function

function CreatePDFfromInvoice(iOrd)
x=OpenRS("P_sel_OrderUIDFromOrNo "&iOrd&","&iDB)
sPageSaveURL=sSiteURL&"/MyInvoices.asp?id="&rsTemp(0)
x=closeRS()
sPageSavePath=sSitePath&"\pdfs\"
x=makepath(sPageSavePath)
x=savePagetoPDF(sPageSaveURL,sPageSavePath,iOrd&"-"&idb)
CreatePDFfromInvoice=GetAttachId(x,iOrd,0)
end function

function GetAttachId(sPath,iOr,iCus)
'x=rwe(iOr)
iOr=cdbl(iOr)
if not iOr=0 then
	sTbl="orders"
	sField="or_number"
	sKey=iOr
end if
if not iCus=0 and iOr=0 then
	sTbl="customers"
	sField="cus_id"
	sKey=iCus
end if
sExt=right(sPath,3)
sSize=GetFileSizeType(sPath)
sFile=GetTextAfterLast(sPath,"\")
sPath=replace(sPath,"\\","\")
sIns="P_Insert_fileDetails '"&sPath&"','"&sSize&"','"&sExt&"','"&sExt&"','"&sFile&"'"
sIns=sIns&",1,'"&sTbl&"','or_number',"&sKey&","&iDB&""
'response.write sIns
x=openRS(sIns)
GetAttachId=RsTemp(0)
x=closeRS()
end function

function GetTextAfterLast(sText,sSplitter)
iStart=1
do until instr(iStart,sText,sSplitter)=0
	iStart=instr(iStart,sText,sSplitter)+1
	GetTextAfterLast=right(sText,len(sText)-iStart+1)
	if iStart>len(sText) then exit function
loop
end function

function savePagetoPDF(sPageSave,sSavePath,sFileName)
Set Session("doc") = Server.CreateObject("ABCpdf7.Doc")
Set theDoc = Session("doc")
'x=ifa(sPageSave)
theID = theDoc.AddImageUrl(sPageSave, True, 0, False)
For i = 1 To 10 ' add up to 3 pages
  If theDoc.GetInfo(theID, "Truncated") <> "1" Then Exit For
  theDoc.Page = theDoc.AddPage()
  theID = theDoc.AddImageToChain(theID)
Next
sFullSave=sSavePath&sFileName
if not right(sFullSave,4)=".pdf" then sFullSave=sFullSave&".pdf"
'x=ifa(sFullSave)
theDoc.save sFullSave
Set Session("doc")=nothing
savePagetoPDF=sFullSave
end function


function EmailGetSysEml(sEmlType,iDB)
'Depending on the staff member executing the script you may want your system emails to come from that staff member
if session("db_send_sys_mail_staff") and not session("staff_email")="" then
	EmailGetSysEml=""""&session("staff_first_name")&" "&session("staff_first_name")&""" <"&session("staff_email")&">"
	EmailGetSysEml=session("staff_email")
	exit function
end if
Select case sEmlType
	case "order_confirm"
		if not len(sSiteEmailSales)=0 then
			EmailGetSysEml=sSiteEmailSales
		else
			EmailGetSysEml=sSiteEmail1
		end if
	case "payment_confirm"
		if not len(sSiteEmailAccounts)=0 then
			EmailGetSysEml=sSiteEmailAccounts
		else
			EmailGetSysEml=sSiteEmail1
		end if		
	case "tracking_confirm"
		if not len(sSiteEmail1)=0 then
			EmailGetSysEml=sSiteEmail1
		else
			EmailGetSysEml="info@botmax.co.nz"
		end if
	case "customer_confirm"
		if not len(sSiteEmail1)=0 then
			EmailGetSysEml=sSiteEmail1
		else
			EmailGetSysEml="info@botmax.co.nz"
		end if
	case "ac_pay1"
		if not len(request("frm"))=0 then
			EmailGetSysEml=request("frm")
		else
			EmailGetSysEml=getDefaultAccountsSend()
		end if
	case else
		if not len(sSiteEmail1)=0 then
			EmailGetSysEml=sSiteEmail1
		else
			EmailGetSysEml="info@botmax.co.nz"
		end if
end Select
end function

function getDefaultAccountsSend()
getDefaultAccountsSend="dan@laptopbattery.co.nz"
end function 



function EmailGetOrderFromId(eid)
if eid>0 then
	'superseed template.
	'check email to see if address is in customers
	x=openRS("P_SEL_Emails_ord_link "&eid)
	if not rsTemp.eof then EmailGetOrderFromId=rsTemp(0)
end if
end function

function EmailGetCustomerFromId(eid)
	'superseed template.
	'check email to see if address is in customers
	x=openRS("P_SEL_Emails_cus_link "&eid)
	if not rsTemp.eof then iCu=rsTemp(0)
EmailGetCustomerFromId=iCu
end function

function EmailGetCustomerFromOrder(iOrd)
if iOrd>0 then
	'superseed template.
	'check email to see if address is in customers
	x=openRS("P_SEL_Cus_from_order "&iOrd&","&idb)
	if not rsTemp.eof then iCu=rsTemp(0)
	x=closeRS()
end if
EmailGetCustomerFromOrder=iCu
end function

function EmailConvertBody(sText,iOr,iCu)
'x=rwb(sText)
'x=rwe(sText)
sOpen="[["
sClose="]]"
x=OpenRS("P_Sel_FieldListByTable 'DB_owner'")
dim sArrayFields(250)
i=0
do until rsTemp.eof
	i=i+1
	'x=rwb(sArrayFields(i)&":"&rsTemp(0))
	sArrayFields(i)=ucase(rsTemp(0))
	rsTemp.movenext
loop
x=closeRS()
x=OpenRS("P_Sel_FieldListByTable 'Orders'")
do until rsTemp.eof
	i=i+1
	sArrayFields(i)=ucase(rsTemp(0))
	rsTemp.movenext
loop
x=closeRS()
x=OpenRS("P_Sel_FieldListByTable 'customers'")
do until rsTemp.eof
	i=i+1
	sArrayFields(i)=ucase(rsTemp(0))
	rsTemp.movenext
loop
x=closeRS()
x=OpenRSA("P_Sel_OrderCusDBOwner "&iOR&","&iCU&","&idb)

for j=1 to i
	ON ERROR RESUME NEXT
	bSkip=0
	sFldVal=rsTempA(sArrayFields(j))
	'x=rwb(sArrayFields(j)&":"&sFldVal)
	if isnull(sFldVal) then sFldVal=""
	if isnumeric(sFldVal) then 
		if sArrayFields(j)="OR_TOTAL" THEN sFldVal=formatNumber(sFldVal,2)
	END IF
	if left(sFldVal,1)="{" then sVal=replaceBrackets(sFldVal)
	sText=replace(sText,sOpen&sArrayFields(j)&sClose,sFldVal,1,-1,1)
	'x=rwb(sOpen&sArrayFields(j)&sClose&":"&sFldVal)
	ON ERROR GOTO 0
next
x=closeRSA()
sText=replace(sText,sOpen&"GETDATE()"&sClose,DateFriendly(now(),1))
EmailConvertBody=sText
'x=rwe(EmailConvertBody)
end function



function CreateArrayFromRS(rsArray)
sArray=""
do until rsArray.eof
	for each field in rsArray.fields
		sArray=sArray&field.name&":"&field.value&";"
	next
	rsArray.movenext
sArray=sArray&"endline;"
loop
CreateArrayFromRS=sArray
end function

function replaceQuotes(sQuoted)
replaceQuotes=replace(replace(sQuoted,"{",""),"}","")
end function

function ItemPhotoAdverts(iStockSys,rs,sLink)
'x=rwe(rs.source)
s=""
'check for adverts that might be linked to this item
set rsA=server.createobject("adodb.recordset")
sSQL="P_Sel_AdvertByID "&iStockSys&","&iDB
'x=ifa(sSQL)
rsA.open sSQL,sMDB
sDeal=rs("dealer3")	
if session("dealer1")=1 then sDeal=rs("dealer1")
if session("dealer2")=1 then sDeal=rs("dealer1")
if session("dealer3")=1 then sDeal=rs("dealer2")	
	
if rsA.eof then
	'use the description build an Advert
	'x=ifa(rs.source)
	'sLink=""
	s=s&"<a id=""imageLink"" href="""&sLink&""" "">Buy "&rs("Cat_Disc_Break3")&"+ of these for only $"&formatNumber(rs("dealer3"),2)&" (each unit).</br></a>"
end if
do until rsA.eof
	sVar=sVar&rs(0)
	rsA.movenext
loop
rsA.close
set rsA=nothing
ItemPhotoAdverts=s
end function


function WhereToLinks()
'popular links that customers might use on a regular basis
s=""
s="<div id=""myRecentItemsContainer""><ul>"
s=s&"<li>"
s=s&"<h3>Where to Next?</h3>"
s=s&"</li>"
s=s&"<li>"
s=s&"<a href=""/categories_index.asp"">Browse categories</a>"
s=s&"</li>"
s=s&"<li>"
s=s&"<a href=""/myPages/myInfo.asp"">My Information</a>"
s=s&"</li>"
s=s&"<li>"
s=s&"<a href=""/MyPages/Favourites.asp"">My Favourites</a>"
s=s&"</li>"
s=s&"<li>"
s=s&"<a href=""/MyPages/RecentlyViewed.asp"">Recently viewed</a>"
s=s&"</li>"
s=s&"<li>"
s=s&"<a href=""/MyPages/myTransactions.asp?Filter=nodispatch"">Orders awaiting dispatch</a>"
s=s&"</li>"
s=s&"<li>"
s=s&"<a href=""/MyPages/myTransactions.asp?Filter=unpaid"">Orders awaiting payment</a>"
s=s&"</li>"
s=s&"<li>"
s=s&"<a href=""/MyPages/myTransactions.asp?Filter=complete"">Completed Orders</a>"
s=s&"</li>"
s=s&"</ul></div>"
WhereToLinks=s
end function

sub HTTPSRedirect()
'if idb=1 and ipAddress<>"127.0.0.1" and ipAddress<>"192.168.1.104" and ipAddress<>"192.168.1.105" then
 	If (Request.ServerVariables("HTTPS") = "off") Then
    x=rda("https://" + Request.ServerVariables("SERVER_NAME")& Request.ServerVariables("PATH_INFO"))
	 end if
'end if
end sub

function errorPage(errType)
'response.redirect ("/errors/error_codes.asp?reason="&errType)
x=rwe("Error: "&errType)
end function

function getAmountFromOrder(iOrderNo,iDBO)
set rs=server.createobject("adodb.recordset")
sSQL="[P_Sel_OrderAmountCus] "&iOrderNo&","&session("cus_sys")&","&iDBO
x=rwe(sSQL)
rs.open sSQL,sMDB
if rs.eof then
	'do something
	response.redirect(errorPage("invalid_order"))
else
	getAmountFromOrder=rs("OR_Total")
end if
rs.close
set rs=nothing
end function



function CreateComboCus(rstDD,iKeyID,ifl_id,sReadonly,sDefault,sFieldValue,sType,bRequired,sfld_script_inline)
strAjax=""
sSQL=""
if sfld_script_inline<>"" then 
	sScript=sScript&" "&sfld_script_inline
	'response.End
end if
if sReadOnly=" readonly=""readonly""" then 
	sReadOnly=" disabled"
	'response.End 
end if
if rstDD("Row_Source_type")="Query" then
	sSQL=ucase(replace(rstDD("Row_Source"),";",""))
	'response.write sSQL&":"&rstDD("tbl_id_link")&"</br>"
	'response.end
	sSelectID="i"&iKeyID&"f"&ifl_id
	sFldName=replace(rstDD("field_friendly_name")," ","_")
	strInput="<select name="""&sFldName&""" "&sScript&" "&sReadOnly&">"
	if bRequired=false then strInput=strInput & "<option value=""""></option>"
	if sSQL<>"" then
		'response.Write(sSQL&"</br>")
		'response.end
		'on error resume next
		set rstLookup=server.createobject("adodb.recordset")
		'on error resume next
		rstLookup.open sSQL,sMDB
		if err.number<>0 then
			'response.Write sSQL
			'response.end				
		end if
		do until rstLookup.eof
			iBound=0	
			strLookupVal=rstLookup(iBound)
			if isnull(sFieldValue) then sFieldValue=""
			sCheck=replace(sFieldValue,"'","")
			sCheck=replace(sCheck," Value=","")
			strInput=strInput & "<option value='" & strLookupVal & "'" 
			if not isnull(strLookupVal) then strLookupVal=cstr(strLookupVal)
			if strLookupVal=sCheck then
				''response.write strLookupVal&":"&sCheck
				'response.end
				strInput=strInput & " Selected"
			end if
			strInput=strInput & ">" & rstLookup(1) & "</option>"
			rstLookup.movenext
		loop
		rstLookup.close
		set rstLookup=nothing
		'response.Write(sSQL&"</br>")
		'response.end
		'on error goto 0
	end if
	'on error goto 0
	strInput=strInput & "</select>"
	'x=rwe(CreateComboCus)
end if
CreateComboCus=strInput
end Function

function DeleteCusData(iTbl,iKey)
'test
if iTbl="0" and ifldID<>"" and ifldID<>"0" then
	iTbl=GetTableID(ifldID)
end if
sDel="P_Del_Record "&iKey&","&iTbl&","&session("dbo")&","&session("cus_sys")&","&Session("Staff_sec_Level")
'x=rwe(sDel)
call OpenDBA()
ConnA.execute(sDel)
ConnA.execute ("P_Ins_UserTableEdit "&session("cus_sys")&","&iTbl&","&iKey&",3"&","&session("dbo"))
call OpenDBA()
response.write "r"&request("k")&":"
'if request("rt")<>"" then response.redirect(request("rt"))
response.end
end function

function ValidateFieldsNew(rsDD,sKey)
'check data if it passes validation then allow save.
dim iErrors
bFieldPass=true
with rsDD
do until.eof
	iFldID=.fields("FL_ID")
	'Check field by Field
	sFieldName=.fields("field_name")
	sValue=request("FC"&iFldID)
	sValidation=.fields("field_type_des")
	sValText=.fields("Validation_Text")
	iMaxLength=.fields("field_length")
	bReq=.fields("Required")
	sPass=false
	if .fields("fl_id")=2035 then 
		sValue=session("cus_user_id")
	end if
	'x=rwe(rsDD.source)
 	'x=rwe("FC"&iFldID&":"&sValue)
 	if rsDD("key_field")=true then iKeyFieldID=rsDD("fl_ID")
	Select case sValidation
		case "bit"
			sPass=true
		case "int"
	if rsDD("field_name")="msg_to" then
		'response.write sPass&sValue
		'response.end
	end if
			if isnumeric(sValue)=true then
				if sValue>-2147483648 and sValue<2147483647 then
					sPass=true					
				else
					sErrCode=sErrCode & "<e"&iFldID&">" & sValText & "</e"&iFldID&">"
					iErrors=iErrors+1
				end if
			else
				if bReq=false and sValue="" then
					sPass=true			
				else
					sErrCode=sErrCode & "<e"&iFldID&">" & sValText & " and is Required</e"&iFldID&">"
					iErrors=iErrors+1
				end if
			end if
		case "smallint"
			if isnumeric(sValue)=true then
				if sValue>-32768 and sValue<32768 then
					sPass=true					
				else
					sErrCode=sErrCode & "<e"&iFldID&">" & sValText & "</e"&iFldID&">"
					iErrors=iErrors+1
				end if
			else
				if bReq=false and sValue="" then
					sPass=true			
				else
					sErrCode=sErrCode & "<e"&iFldID&">" & sValText & " and is Required</e"&iFldID&">"
					iErrors=iErrors+1
				end if
			end if
		case "smallint"
			if isnumeric(sValue)=true then
				if sValue>-1 and sValue<256 then
					sPass=true					
				else
					sErrCode=sErrCode & "<e"&iFldID&">" & sValText & "</e"&iFldID&">"
					iErrors=iErrors+1
				end if
			else
				if bReq=false and sValue="" then
					sPass=true			
				else
					sErrCode=sErrCode & "<e"&iFldID&">" & sValText & " and is Required</e"&iFldID&">"
					iErrors=iErrors+1
				end if
			end if	
		case "date","datetime","smalldatetime"
			if isdate(sValue)=true then
				sPass=true				
			else
				if bReq=false and sValue="" then
					sPass=true
				else
					sErrCode=sErrCode & "<e"&iFldID&">This field must have a date value.</e"&iFldID&">"
					iErrors=iErrors+1
				end if
			end if
		case else 
			if len(sValue)>iMaxLength then
				sPass=false
				sErrCode=sErrCode & "<e"&iFldID&">You have entered to much text.  Maximum allowed is "&iMaxLength&". Please reduce by "&abs(iMaxLength-len(sValue))&" characters.</e"&iFldID&">"
			else
				if bReq=true and (sValue="" or sValue="</br>") then
					sErrCode=sErrCode & "<e"&iFldID&">This field must have a value.</e"&iFldID&">"
					iErrors=iErrors+1
				else
					sPass=true			
				end if		
			end if
	end select

	if sPass=false then bFieldPass=false
	if sPass=false then 
			'response.write "No Pass for :" & sFieldName & "=" & sValue &"</br>"
			'response.end
	end if
	.movenext
loop
rsDD.movefirst 
'response.write sPassVal
'response.end
end with
ValidateFieldsNew=sErrCode
if ValidateFieldsNew<>"" then 
	ValidateFieldsNew="<validate><failed><table class=""tResultsFail""><tr><td><img src=""/styles/images/error.gif""></td><td><span class=""error"" />Please Check the form "&iErrors&" errors were found.</span></td></tr></table></failed>"&ValidateFieldsNew&"</validate>"
else 
	ValidateFieldsNew=""
end if
end function

function saveCusData(iField,iTbl,iGroup,iAjax)
sDBO=""

sKey=stripinvalid(request("key"))
if sKey="" then sKey=0
'Allows edit of cusData via field,table, optional group limiter use 0 for all groups
set rstDD=server.createobject("adodb.recordset")
sSQl="P_Sel_FieldLookup_CusEdit "&iField&","&iTbl&","&iGroup
'x=rwe(sSQL)
rstDD.open sSQL,sMDB
if rstDD.eof then
	response.write "EOF"
	response.end
end if
'now with rs getData
set rs=server.createobject("adodb.recordset")
sValidated=ValidateFieldsNew(rstDD,sKey)
if len(sValidated)>0 then 
	'must be an error
	response.write (sValidated)
	response.end 
end if
if rstDD("has_db_owner") then	sDBO=session("dbo")
if request("Action")="1" then
	'insert action
	sSQL="P_T_INS_"&rstDD("table_name")&"_By_CUS_SYS "
	sSQL=sSQL&session("cus_sys")&","
else
	'update action
	sSQL="P_T_UP_"&rstDD("table_name")&"_By_CUS_SYS "
	sSQL=sSQL&session("cus_sys")&","&sKey&","
end if
do until rstDD.eof
	sVal=request("FC"&rstDD("fl_id"))
	if rstDD("fl_id")=2035 then 
		sVal=session("cus_user_id")
	end if
	
	sVal=replace(sVal,"'","''")
	if not (request("Action")="1" and sVal="" and len(rstDD("auto_populate"))>0) and not rstDD("key_field") then
		if rstdd("field_type_des")="bit" then
			if sVal="true" then 
				sVal=1
			else
				sVal=0
			end if
		end if
		sQuote="'"
		if instr(rstDD("Field_Type_Des"),"int")>0 then sQuote=""
		sSQL=sSQL&sQuote&sVal&sQuote&","
	else
		'get data from autoPopulate
	end if
	'x=rwb("i"&sKey&"f"&rstDD("fl_id")&":"&sVAL)
	rstDD.movenext
loop
'
'get all fields
sSQL=sSQL&sDBO
if right(sSQL,1)="," then sSQL=left(sSQL,len(sSQL)-1)
	'x=rwe(sSQL)
'on error resume next
call opendb()
conn.execute (sSQL)
call closedb()
if iAjax="1" then
	if iTbl=159 then
		if request("Action")="1" then
			call openRS("P_Sel_StockFavouritesByNoteID "&session("cus_sys")&",0,"&session("dbo"))
			'x=rwe(rsTemp("related_key"))
			s=s&ShowNoteTextByRS(rsTemp)
			call closeRS()
		else
			'call openRS("P_Sel_StockFavouritesByNoteID "&session("cus_sys")&",'"&request("FD1882")&"',"&session("dbo"))
			s=request("FC1969")
		end if
		'x=rwe("P_Sel_StockFavouritesByNoteID "&session("cus_sys")&","&session("dbo"))
		x=rw(s)
		x=rwe("")
		'get note text data
	end if
else
	if err.number=0 then
		x=rwe(GetSavedLinks())
	else
		x=rwe(GetErrorOccured(sSQL))
	end if
end if
on error goto 0
end function

function openRSPhoto(SQL)
set rsPhoto=server.createObject("adodb.recordset")
'z=rwb(SQL)
rsPhoto.open SQL,sMDB
end function

function closeRSPhoto()
rsPhoto.close
end function

function openRSBig(SQL)
dim rsBig
set rsBig=Server.CreateObject("ADODB.recordset")
set conn=Server.CreateObject("ADODB.Connection")
conn.Connectionstring=sMDB
conn.CommandTimeout=3000
conn.open
rsBig.Open sSQL, conn
conn.close
end function

function closeRSBig()
on error resume next
If Not rsBig Is Nothing Then
	rsBig.close()
	conn.close
	set conn=nothing
	set rsBig=nothing
end if
on error goto 0
end function

function runUpdate(SQL)
call openDB()
conn.execute(SQL)
call closeDB()
set CnTax=nothing
end function

function openRS(SQL)
set rsTemp=server.createObject("adodb.recordset")
'z=rwb(SQL)
rsTemp.open SQL,sMDB
end function


function openRSA(SQL)
set rsTempA=server.createObject("adodb.recordset")
'z=rwb(sMDBA)
'x=rwb(SQL)
rsTempA.open SQL,sMDBA

end function

function closeRSA()
on error resume next
If Not rsTempA Is Nothing Then
	rsTempA.close()
	set rsTempA=nothing
end if
on error goto 0
end function



function insRSTax(SQL)
set CnTax=server.createobject("ADODB.connection")
CnTax.Connectionstring=sTaxDB
CnTax.open
CnTax.execute(SQL)
CnTax.close
set CnTax=nothing
end function

function openRSTax(SQL)
'z=rwe(SQL)
set rsTempTax=server.createObject("adodb.recordset")
'x=rwe(sTaxDB)
rsTempTax.open SQL,sTaxDB
'x=rwb(rsTempTax.source)
end function

function closeRSTax()
rsTempTax.close
set rsTempTax=nothing
end function

function openRSTax2(SQL)
'z=rwe(SQL)
set rsTax=server.createObject("adodb.recordset")
'x=rwe(sTaxDB)
rsTax.open SQL,sTaxDB
'x=rwb(rsTempTax.source)
end function

function closeRSTax2()
rsTax.close
set rsTax=nothing
end function

function openRSM(SQL)
set rsTemp=server.createObject("adodb.recordset")
rsTemp.open SQL,sMailerDB
end function

function closeRSM(SQL)
set rsTemp=server.createObject("adodb.recordset")
rsTemp.open SQL,sMailerDB
end function

function closeRS()
If Not rsTemp Is Nothing Then
	on error resume next
	rsTemp.close()
	on error goto 0
	set rsTemp=nothing
end if
end function

function GetSavedLinks()
s=""
s=sHTMSuccess
s=s&"<span class=""passed"">Save was sucessfull!</span>"
GetSavedLinks=s
end function

function GetErrorOccured(sSQL)
s=""
s="<span class=""attention"">Error Occured with this Please try again.</br></br>"
x=ifa(sSQL)
s=s&"</span>"
GetErrorOccured=s
end function

function getCusDataForEdit(iField,iTbl,iGroup)
'Allows edit of cusData via field,table, optional group limiter use 0 for all groups
set rstDD=server.createobject("adodb.recordset")
'x=rwe(sPage)
sSQl="P_Sel_FieldLookup_CusEdit "&iField&","&iTbl&","&iGroup
'x=rwe(sSQL)
rstDD.open sSQL,sMDB
if rstDD.eof then
	response.write "EOF"
	response.end
end if
'now with rs getData
set rs=server.createobject("adodb.recordset")
sSQL="P_T_Sel_"&rstDD("table_name")&"_By_CUS_SYS "&session("cus_sys")
if request("id")<>"" then	
	sSQL=sSQL&",'"&request("id")&"'"
	iKeyID=request("id")
else
	if iTbl=13 then
		sSQL=sSQL&""
		iKeyID=session("cus_sys")
	else
		sSQL=sSQL&",0"
		iKeyID=0
	end if
end if
if rstDD("has_db_owner") then	sSQL=sSQL&","&idb
rs.open sSQL,sMDB
'x=rwe(sSQL)
if rs.eof then
	sFrmType="Add"
	sHTMAdd="<input type=""hidden"" name=""Action"" value=""1"">"
else
	sFrmType="Edit"
	sHTMAdd="<input type=""hidden"" name=""Action"" value=""0"">"
	if request("id")="" then 
		ikey=0
	else
		ikey=stripinvalid(request("id"))
	end if
	sHTMAdd=sHTMAdd&"<input type=""hidden"" name=""key"" value="""&ikey&""">"
end if
sPage=sScriptname
sPage=replace(sPage,"Edit.asp","")
sPage=replace(sPage,"Add.asp","")
sPage="/MyPages/"&sPage&"Save.asp"
'x=rwe(sPage)
sHTMLFld="<form name="""&replace(sScriptname,".asp","")&""" action="""&sPage&""" method=""post"">"
sHTMLFld=sHTMLFld&sHTMAdd
sHTMLFld=sHTMLFld&"<table class=""cusData"">"
'ok sus out is data for edit or add.
'x=rwb(rs.source)
with rs
   do until rstDD.eof
	  'First Check to see if the Data for the page is returned from an error
		if rstDD("fld_Script_path")<>"" then 
			sScripts=""
			sScripts="<script language=""javascript"" type=""text/javascript"" src="""&rstDD("fld_Script_path")&"""></script>"
		end if
		sGroupHTML=""
    ifl_id=rstDD("FL_ID")
    sFldName=rstDD("Field_Name")
    iOrder=rstDD("Field_Order")
    iID=rstDD("FL_ID")
    bKey=rstDD("Key_field")
    sType=rstDD("Field_Type_Des")
    strValClass="regular"
    sValText=""
    strInputClass=""
    sAddedCode=""
    sReadonly=""
  	if rstDD ("fld_search")=true then strInputClass="fldSearch" 
   	'x=rwb(sFldName)
  	'response.end
    if bKey then strInputClass="fldkey"
    if not .eof then
    	sFieldValue=.fields(sFldName)
  	end if
    if rstDD("IS_Shown")=true then
    	'CREATE THE GROUP NAME HEADING//////////////////////////////////////////////
	    if sGroupName<>rstDD("group_name") then
				iGroupCount=iGroupCount+1
				sGroupName=rstDD("group_name")
				iGroupID=rstDD("fld_group_id")
			if iGroupCount>1 then sRowSeperator="<tr><td colspan=""2"" class=""trSep""></br> </td></tr>" else sRowSeperator=""
				sGroupHTML=sRowSeperator&"<tr><td colspan=""2""><div class=""formGroup""><span class=""formGroupLeft""> "&rstDD("group_name")&"</span></div></td></tr>"
			end if
	    sHTMLFld=sHTMLFld&sGroupHTML&"<tr>"
	    '////////////////////////////////////////////////////////////////////////////
	    bEditSec=1
    	if rstDD("IS_Shown")=true then
				sHTMLFld=sHTMLFld&"<td class=""fldName"">"&rstDD("Field_Friendly_name")&"</td>"
				'x=RWS("fieldAuctual:"&fID(iFlds)&"="&fEdit(iFlds)&":"&iFlds&"</br>")
				sReadOnly=""
			  if request("err_code")<>"" then
					for i=0 to Ubound(PassArray)
						if ifl_id & "_pass" =  PassArray(i) then
							sFieldValue=PassArray(i+1)
							exit for
						end if
					next
					for i=0 to Ubound(ErrorArray)
						if ifl_id  =  ErrorArray(i) then
							sValText=" ** " & ErrorArray(i+1)
							strValClass="highlight_error"
						end if
						if ifl_id  & "_val" =  ErrorArray(i) then
							sFieldValue = ErrorArray(i+1)
						end if								
					next
				end if
				on error goto 0
				if sFieldValue="" and rstDD("Auto_Populate")<>"" and iKeyID=0 then
					sFieldValue=rstDD("Auto_Populate")
					'Note the enhanced version enables administrators to add a default vaulue using a procedure or SQL free text
					'Also note this can also include the session and page variables as required.
					'Note format format for the default value should of the type =1 or  =(Select getdate()) or ="default text"
					'Further note Default Value is only applied when a record is being added and has no affect at other times.
					sAuto=rstDD("Auto_Populate")
					'on error resume next
					if left (sAuto,1)="=" then sAuto=right(sAuto,len(sAuto)-1)
					Select case sType
					case "bit" 
						if isnull(sAuto) then sAuto=""
						sAuto=replace(sAuto,"=","")
						'x=rwe(sAuto)
						sBit=replaceBrackets(sAuto)
						sFieldValue=cBool(sBit)
					case "float" ,"int" ,"int" ,"money" ,"real" ,"smallint"
						if instr(sAuto,"=")>0 then
							if not isnumeric(sAuto) then 
								response.write "<span class=""error"">Please note parameter "&rstDD("Field_Name")&" default value is invalid for datatype "&sType&".</span>"
								if iSec>9 then response.write "<a href=""/admin/item.asp?t=19&id="&rstDD("fl_id")&""" >Edit</a>" 
							else
								sFieldValue=Cint(replace(sAuto,"=",""))
							end if
						else
							sFieldValue=sAuto
						end if
					case 	"datetime","date"
						sFieldValue=replaceBrackets(sFieldValue)
						'x=rwe(sFieldValue)
						if sFieldValue="date" or sFieldValue="now" or sFieldValue="getdate" then
							sFieldValue=SQLDate(now(),true)
						else
							sFieldValue="<span class=""error"">Please note parameter "&rstDD("fl_id")&" default value is invalid for datatype "&sType&".</span>"
						end if
					case "varchar" ,"nvarchar" ,"nchar"
						sFieldValue=sAuto
						sDefault=sFieldValue
					case else
						sFieldValue="unknown field type"
					end select
				end if
				on error goto 0
				if request("fd"&rstDD(0))<>"" then 
					sFieldValue=request("fd"&rstDD(0))
					'response.write sFieldValue
					'response.End
				end if
				if iCpyKeyID<>"" and  sFieldValue="" then
					sFieldValue=.fields(sFldName)
				end if
				if CreateDropDown="" then
					if ifl_id=91 then 
						'x=rwe("readonly:"&sReadOnly)
					end if
					strInput=Converttohtml_Input(replace(rstDD("Field_Friendly_Name")," ","_"),sFieldValue,rstDD("FIELD_TYPE_DES"),rstDD("required"),rstDD("FIELD_LENGTH"),rstDD("fld_script_inline"),sReadOnly,strInputClass,rstDD("IS_Shown"),rstDD("File_Upload"),strKey,iKeyID,ifl_id,rstDD ("fld_search"),bAutoSave)
					'if sType="bit" then x=rwe(strInput)
				else
					strInput=CreateDropDown
				end if
				on error goto 0
				if rstDD("required")=true then
					strValidation="<span align=""right"" class=""attention"">*</span>"
				else
					strValidation=""
				end if
				if sValText<>"" then strValidation=sValText
			end if
			if not isnull(rstDD("Row_Source_type")) and rstDD("has_combo")=true then 
				strInput=CreateComboCus(rstDD,iKeyID,ifl_id,sReadOnly,sDefault,sFieldValue,rstDD("FIELD_TYPE_DES"),rstDD("required"),rstDD("fld_script_inline"))
			end if
			if rstDD("IS_Shown")=true then
				if iSec<7 then
					sToolTip=rstDD("field_description")
				else
					sToolTip=rstDD("field_description")
				end if
				sHTMLFld=sHTMLFld&"<td class=""fldVal""><span>"&strInput&" "& strValidation&""
				if len(rstDD("field_description"))>0 THEN
					sTipWidth=cint(len(sToolTip)/3+100) &"px"
					'sDiv="<div class=""fldHelp"" id=""dh"&ifl_id&""" onMouseover=""clearhidetip()"" onMouseout=""delayhidetip('dh"&ifl_id&"',800)"">"
					'Asumption is made that for adding help files the first "." indicates header row is finished.
					'If no "." is found then or positionb of "." is near the end more than 60% fo the length there is no title.
					if iSec>7 then
						sT=rstDD("Field_Name")
					else
						sT=rstDD("Field_Friendly_Name")
					end if
					sB=sToolTip
					sHTMTbl="<table class=""tblLight""><tr class=""dhHead""><td>"&sT&"</td></tr>"
					sHTMTbl=sHTMTbl&"<tr class=""dhBody""><td>"&replace(replace(sB,chr(13),"</br>"),chr(10),"</br>")&"</td></tr></table>"
					sDiv=sDiv&sHTMTbl
					sDiv=sDiv&"</div>"&vbcrlf
					'sHTMLFld=sHTMLFld&"<img class=""fldinfoimg"" src=""/Styles/Images/info.png"" onMouseover=""fixedtooltip('dh"&ifl_id&"', this, event, 'auto')"" onMouseout=""delayhidetip('dh"&ifl_id&"',800)"">"&sDiv
					sDiv=""
				end if
				sHTMLFld=sHTMLFld&strAjax&"</span>"
				sHTMLFld=sHTMLFld&"<span id=""E"&ifl_id&""" class=""error""></span>"
				'comment field groups are show here.
	    	'sHTMLFld=sHTMLFld&"<td><img class=""ajaxresponse"&ifl_id&""" src=""""/></td>"
			else
				if rstDD("Field_Name")<>"DB_Owner" then
					sHTMLFld=sHTMLFld&strInput 
				end if
			end if
			k=k+1
		end if
		'x=rwb(sHTMLFld)
		'sHTMLFld=""
		rstDD.movenext
		strInput=""
		strValidation=""
		strValClass=""
		sFieldValue=""
		strAjax=""
		CreateDropDown=""
		sHTMLFld=sHTMLFld&"</tr>"&vbcr
		iFlds=iFlds+1
 loop
end with
rstdd.close
set rs=nothing
'add save button
sHTMLFld=sHTMLFld&"<tr><td colspan=""6"" align=""center""><input type=""submit"" value=""Save Changes""></td></tr>"
sHTMLFld=sHTMLFld&"</table></form>"
getCusDataForEdit=sHTMLFld
end function

function replaceBrackets(sText)
replaceBrackets=replace(replace(sText,"(",""),")","")
replaceBrackets=replace(replace(replaceBrackets,"{",""),"}","")
end function


function getCusDataForEditAjax(iField,iTbl,iGroup)
'Allows edit of cusData via field,table, optional group limiter use 0 for all groups
set rstDD=server.createobject("adodb.recordset")
sSQl="P_Sel_FieldLookup_CusEdit "&iField&","&iTbl&","&iGroup
'x=rwe(sSQL)
rstDD.open sSQL,sMDB
if rstDD.eof then
	response.write "EOF"
	response.end
end if
'now with rs getData
set rs=server.createobject("adodb.recordset")
sSQL="P_T_Sel_"&rstDD("table_name")&"_By_CUS_SYS "&session("cus_sys")
if request("id")<>"" then	
	sSQL=sSQL&",'"&request("id")&"'"
	iKeyID=request("id")
else
	sSQL=sSQL&",0"
	iKeyID=0
end if
if rstDD("has_db_owner") then	sSQL=sSQL&","&idb
rs.open sSQL,sMDB
'x=rwe(sMDB)
if rs.eof then
	sFrmType="Add"
	sHTMAdd="<input type=""hidden"" name=""Action"" value=""1"">"
else
	sFrmType="Edit"
	sHTMAdd="<input type=""hidden"" name=""Action"" value=""0"">"
	if request("id")="" then 
		ikey=0
	else
		ikey=stripinvalid(request("id"))
	end if
	sHTMAdd=sHTMAdd&"<input type=""hidden"" name=""key"" value="""&ikey&""">"
end if

sPage="/MyPages/"&replace(sScriptname,".asp","")&"Save.asp"
'x=rwe(sSQL)
sHTMLFld="<form name="""&replace(sScriptname,".asp","")&""" action="""&sPage&""" method=""post"">"
sHTMLFld=sHTMLFld&sHTMAdd
sHTMLFld=sHTMLFld&"<table class=""cusData"">"
'ok sus out is data for edit or add.

with rs
   do until rstDD.eof
	  'First Check to see if the Data for the page is returned from an error
		if rstDD("fld_Script_path")<>"" then 
			sScripts=""
			sScripts="<script language=""javascript"" type=""text/javascript"" src="""&rstDD("fld_Script_path")&"""></script>"
		end if
		sGroupHTML=""
    ifl_id=rstDD("FL_ID")
    sFldName=rstDD("Field_Name")
    iOrder=rstDD("Field_Order")
    iID=rstDD("FL_ID")
    bKey=rstDD("Key_field")
    sType=rstDD("Field_Type_Des")
    strValClass="regular"
    sValText=""
    strInputClass=""
    sAddedCode=""
    sReadonly=""
  	if rstDD ("fld_search")=true then strInputClass="fldSearch" 
   	'response.write rstDD.Source
  	'response.end
    if bKey then strInputClass="fldkey"
    if not .eof then
    	sFieldValue=.fields(sFldName)
  	end if
    if rstDD("IS_Shown")=true then
    	'CREATE THE GROUP NAME HEADING//////////////////////////////////////////////
	    if sGroupName<>rstDD("group_name") then
				iGroupCount=iGroupCount+1
				sGroupName=rstDD("group_name")
				iGroupID=rstDD("fld_group_id")
			if iGroupCount>1 then sRowSeperator="<tr><td colspan=""2"" class=""trSep""></br> </td></tr>" else sRowSeperator=""
				sGroupHTML=sRowSeperator&"<tr><td colspan=""2""><div class=""formGroup""><span class=""formGroupLeft""> "&rstDD("group_name")&"</span></div></td></tr>"
			end if
	    sHTMLFld=sHTMLFld&sGroupHTML&"<tr>"
	    '////////////////////////////////////////////////////////////////////////////
	    bEditSec=1
    	if rstDD("IS_Shown")=true then
				sHTMLFld=sHTMLFld&"<td class=""fldName"">"&rstDD("Field_Friendly_name")&"</td>"
		  end if
			'x=RWS("fieldAuctual:"&fID(iFlds)&"="&fEdit(iFlds)&":"&iFlds&"</br>")
			sReadOnly=""
		  if request("err_code")<>"" then
				for i=0 to Ubound(PassArray)
					if ifl_id & "_pass" =  PassArray(i) then
						sFieldValue=PassArray(i+1)
						exit for
					end if
				next
				for i=0 to Ubound(ErrorArray)
					if ifl_id  =  ErrorArray(i) then
						sValText=" ** " & ErrorArray(i+1)
						strValClass="highlight_error"
					end if
					if ifl_id  & "_val" =  ErrorArray(i) then
						sFieldValue = ErrorArray(i+1)
					end if								
				next
			end if
			on error goto 0
			if sFieldValue="" and rstDD("Auto_Populate")<>"" and iKeyID=0 then
				sFieldValue=rstDD("Auto_Populate")
				'Note the enhanced version enables administrators to add a default vaulue using a procedure or SQL free text
				'Also note this can also include the session and page variables as required.
				'Note format format for the default value should of the type =1 or  =(Select getdate()) or ="default text"
				'Further note Default Value is only applied when a record is being added and has no affect at other times.
				sType=rstDD("Field_Type_Des")
				sAuto=rstDD("Auto_Populate")
				'on error resume next
				if left (sAuto,2)="=(" then
					'must be evaluation of a SQL term
					'Two posibilities 1 field or Two
					SDefaultText=""
					set rs=server.createobject("adodb.recordset")
					sAuto=trim(sAuto)
					sqlEval=mid(sAuto,3,len(sAuto)-3)
					if right(sAuto,1)<>")" then
						response.write "incorrect SQL format please correct with closing <b>)</a> bracket.</br> " 
					else
						sqlEval=mid(sqlEval,1,len(sqlEval))
					end if
					sqlEval=ReplaceVar(sqlEval,0,0,0,"")	
					if session("debug")=true then x=rwSQL(sqlEval)
					'response.end
					'response.end
					on error resume next
					'response.write sqlEval
					rs.open sqlEval,sMDB
					if err.number<>0 then
						sFieldValue="error with default value SQL please check"
						'response.end
					end if
					iFelds=rs.fields.count
					if not rs.eof then
						if iFelds=1 then
							sFieldValue=rs(0)
						else
							if rstDD("readonly")=true then sdis="disabled=""disabled"""
							CreateDropDown="<select name=""i"&iKeyID&"f"&ifl_id&""" id=""i"&iKeyID&"f"&ifl_id&""" "&sdis&"><option value="""&rs(0)&""">"&rs(1)&"</option></Select>"
							sdis=""
						end if
					end if
					rs.Close
					set rs=Nothing
				else
					if sType="bit" then sFieldValue=cBool(replace(sAuto,"=",""))
					if (sType="float" or sType="int" or sType="int" or sType="money" or sType="real" or sType="smallint") then 
						if instr(sAuto,"=")>0 then
							if not isnumeric(sAuto) then 
								response.write "<span class=""error"">Please note parameter "&rstDD("Field_Name")&" default value is invalid for datatype "&sType&".</span>"
								if iSec>9 then response.write "<a href=""/admin/item.asp?t=19&id="&rstDD("fl_id")&""" >Edit</a>" 
							else
								sFieldValue=Cint(replace(sAuto,"=",""))
							end if
						else
							sFieldValue=sAuto
						end if
					end if
					on error goto 0
					if sType="datetime" then 
						if sFieldValue="=date()" or sFieldValue="date()" or sFieldValue="=now()" or sFieldValue="now()" then
							sFieldValue=CvbShortdateTime(now(),true)
						else
							sFieldValue="<span class=""error"">Please note parameter "&rstDD("field_id")&" default value is invalid for datatype "&sType&".</span>"
						end if
					end if
					'on error goto 0
					if (sType="varchar" or sType="nvarchar" or sType="nchar") then sFieldValue=sAuto
					sDefault=sFieldValue
				end if
			end if
			on error goto 0
			if request("fd"&rstDD(0))<>"" then 
				sFieldValue=request("fd"&rstDD(0))
				'response.write sFieldValue
				'response.End
			end if
			if iCpyKeyID<>"" and  sFieldValue="" then
				sFieldValue=.fields(sFldName)
			end if
			if CreateDropDown="" then
				if ifl_id=91 then 
					'x=rwe("readonly:"&sReadOnly)
				end if
				strInput=Converttohtml_Input("i"&iKeyID&"f"&ifl_id,sFieldValue,rstDD("FIELD_TYPE_DES"),rstDD("required"),rstDD("FIELD_LENGTH"),rstDD("fld_script_inline"),sReadOnly,strInputClass,rstDD("IS_Shown"),rstDD("File_Upload"),strKey,iKeyID,ifl_id,rstDD ("fld_search"),bAutoSave)
			else
				strInput=CreateDropDown
			end if
			on error goto 0
			if rstDD("required")=true then
				strValidation="<span align=""right"" class=""attention"">*</span>"
			else
				strValidation=""
			end if
			if sValText<>"" then strValidation=sValText
		end if
		if not isnull(rstDD("Row_Source_type")) and rstDD("has_combo")=true then 
			strInput=CreateComboCus(rstDD,iKeyID,ifl_id,sReadOnly,sDefault,sFieldValue,rstDD("FIELD_TYPE_DES"),rstDD("required"),rstDD("fld_script_inline"))
		end if
		if rstDD("IS_Shown")=true then
			if iSec<7 then
				sToolTip=rstDD("field_description")
			else
				sToolTip=rstDD("field_description")
			end if
			sHTMLFld=sHTMLFld&"<td class=""fldVal""><span>"&strInput&" "& strValidation&""
			if len(rstDD("field_description"))>0 THEN
				sTipWidth=cint(len(sToolTip)/3+100) &"px"
				sDiv="<div class=""fldHelp"" id=""dh"&ifl_id&""" onMouseover=""clearhidetip()"" onMouseout=""delayhidetip('dh"&ifl_id&"',800)"">"
				'Asumption is made that for adding help files the first "." indicates header row is finished.
				'If no "." is found then or positionb of "." is near the end more than 60% fo the length there is no title.
				if iSec>7 then
					sT=rstDD("Field_Name")
				else
					sT=rstDD("Field_Friendly_Name")
				end if
				sB=sToolTip
				sHTMTbl="<table class=""tblLight""><tr class=""dhHead""><td>"&sT&"</td></tr>"
				sHTMTbl=sHTMTbl&"<tr class=""dhBody""><td>"&replace(replace(sB,chr(13),"</br>"),chr(10),"</br>")&"</td></tr></table>"
				sDiv=sDiv&sHTMTbl
				sDiv=sDiv&"</div>"&vbcrlf
				sHTMLFld=sHTMLFld&"<img class=""fldinfoimg"" src=""/Styles/Images/info.png"" onMouseover=""fixedtooltip('dh"&ifl_id&"', this, event, 'auto')"" onMouseout=""delayhidetip('dh"&ifl_id&"',800)"">"&sDiv
				sDiv=""
			end if
			sHTMLFld=sHTMLFld&strAjax&"</span>"
			sHTMLFld=sHTMLFld&"<span id=""E"&ifl_id&""" class=""error""></span>"
			'comment field groups are show here.
    	sHTMLFld=sHTMLFld&"<td><img class=""ajaxresponse"&ifl_id&""" src=""""/></td>"
		else
			if rstDD("Field_Name")<>"DB_Owner" then
				sHTMLFld=sHTMLFld&strInput 
			end if
		end if
		k=k+1
		rstDD.movenext
		strInput=""
		strValidation=""
		strValClass=""
		strAjax=""
		CreateDropDown=""
		sHTMLFld=sHTMLFld&"</tr>"&vbcr
		iFlds=iFlds+1
 loop
end with
rs.close
set rs=nothing
'add save button
sHTMLFld=sHTMLFld&"<tr><td colspan=""6"" align=""center""><input type=""submit"" value=""Save Changes""></td></tr>"
sHTMLFld=sHTMLFld&"</table></form>"
getCusDataForEditAjax=sHTMLFld

end function

function parseMailID(sMailText)
	if instr(sMailText,"?c=")>0 then
		parseMailID=fnFindText(sMailText,"?c=",3,""">",0,"")
	end if
end function

function MoveAttachmentToSent(iAttachID,sResponse)
if iAttachID<>0 then
	Dim fso, f, fc, nf
	'on error resume next
  Set fso = CreateObject("Scripting.FileSystemObject")
	set rs=server.createobject("adodb.recordset")
	iEmlID=parseMailID(sResponse)
	if iEmlID>0 then
		sSQL="P_Sel_Temp_Files "&iAttachID
		'response.write sSQL&"</br>"
		'lookup the last email with the iAccountID
		rs.open sSQL,sMDB
		do until rs.eof
			iID=rs(0)
			sPath=rs(1)
			sNewPath=sSitePath&"\DB_Owner_files\DB"&idb&"\emails\"&session("eAccount")&"\"&iEmlID&"\"
			if fso.FolderExists (sNewPath)=false then
				'response.write sSitePath&"\DB_Owner_files\DB"&idb&"\emails\"&session("eAccount")&"\"
				call AddNewFolder(sSitePath&"\DB_Owner_files\DB"&idb&"\emails\"&session("eAccount")&"\",iEmlID)
			end if
			'response.write sNewPath&"</br> copying from "&sPath&"</br>"
			fso.CopyFile sPath,sNewPath
			fso.deleteFile sPath
			sSQL="P_Up_Temp_Files '"&sNewPath&"',"&iEmlID&","&iID
			conn.execute (sSQL)
			'response.write "</br>"&sSQL&"</br>"
			'response.end
			rs.movenext
		loop
		rs.close
	end if
	'Delete temp emlID
	call OpenDBA()
	connA.execute("P_Del_Temp_mail "&session("staff_id")&",'random',"&iDB)
	call CloseDBA()
	session("emlTempID")=""
	set rs=nothing
	'response.end
end if
end function


function GetFileSizeType(sFile)
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	'file = Server.MapPath(WhatFile)
	set f = fs.GetFile(sFile)
	intSizeK =f.Size/1
	if intSizeK = 0 then intSizeK = 1
	GetFileSizeType = intSizeK
	set f=nothing
	set fs=nothing
end function

function Write_CSV_From_Recordset( RS )
s=""

if RS.EOF then
  exit function
end if

dim RX
set RX = new RegExp
    RX.Pattern = "\r|\n|,|"""

dim i
dim Field
dim Separator

'
' Writing the header row (header row contains field names)
'

Separator = ""
for i = 0 to RS.Fields.Count - 1
  Field = RS.Fields( i ).Name
  if RX.Test( Field ) then
    '
    ' According to recommendations:
    ' - Fields that contain CR/LF, Comma or Double-quote should be enclosed in double-quotes
    ' - Double-quote itself must be escaped by preceeding with another double-quote
    '
    Field = """" & Replace( Field, """", """""" ) & """"
  end if
  s=s&Separator & Field
  Separator = ","
next
s=s&vbNewLine

'
' Writing the data rows
'

do until RS.EOF
  Separator = ""
  for i = 0 to RS.Fields.Count - 1
    '
    ' Note the concatenation with empty string below
    ' This assures that NULL values are converted to empty string
    '
    Field = RS.Fields( i ).Value & ""
    if RX.Test( Field ) then
      Field = """" & Replace( Field, """", """""" ) & """"
    end if
    s=s&Separator & Field
    Separator = ","
  next
  s=s&vbNewLine
  RS.MoveNext
loop
Write_CSV_From_Recordset=s
end function

function advertise(sAddGroup)
sA="You may be interested in this...</br>"
Select case sAddGroup
	case "myOrders"
		sA=sA&"<span class=""dev_comment"">Dev Team note.  Add items of interest. Use item creator procedure to show items.</span>"
	end select
advertise=sA
end function


Function CreateUserFields(ifldID,iKey,sVal,iSec,iTabOrder)

sLink=""
sLinkC=""
sClass="Text_Class"
if isnull(sVal) then sVal=""
'if len(sVal)>70 then sVal=left(sVal,67)&"..."
if not isnull(sVal) then
	if isNumeric(sVal) then 
		sClass="Number_Class"
		'iTot(x)=iTot(x)+rs.fields(fld.name)
	end if
end if
'on error resume next
sFunction=""
'if session("staff_id")=1 then response.write ifldID&":"&fEdit(iFlds)&"</br>"
'check_fld_sec
'if fEdit(iFlds)<=iSec THEN 
	'response.Write("ifldID("&iFlds&"):"&fID(iFlds)&"</br>")
	if len(ifldID)>0 and iSec<=session("staff_sec_level") then sFunction="uffSingle("&iKey&","&ifldID&")"
'end if
CreateUserFields="<td onclick="""&sFunction&""" tabindex="""&iTabOrder&""" id=""r"&iKey&"f"&ifldID&""" class="""&sClass&"""><div style=""overflow:hidden;height:20px;"">"&sLink&replace(sVal,"""","""""")&sLinkC&"</div></td>"&vbCR
	'sLink="<a href=""item.asp?t="&iTable&"&id="&rs.fields(fld.name)&""">"
	'sLinkC="</a>"
	's=s&"<td tabindex="""&iTabOrder&""" id=""r"&iKey&"f"&ifldID&""" class="""&sClass&"""><div style=""overflow:hidden;height:20px;"">"&sLink&replace(sVal,"""","""""")&sLinkC&"</div></td>"&vbCR
'CreateUserFields=s
end function




function DeliveryAddresses(rsDA)

end function

function FnCreateCombo(fnRST,sName,iBoundColumn,sValueColumns,sJavaScript)
iBoundColumn=cint(iBoundColumn)
s=""
s="<select name="""&sName&""" id="""&sName&""" "&sJavaScript&">"&vbcr
if fnRST.eof then
	s=s&"<option value=""""></option></select>"
	FnCreateCombo=s
	exit function
end if
do until fnRST.eof
	sValues=split(sValueColumns,",")
	sValueDisplayed=""
	
	for i=0 to ubound(sValues)
		iIndexColumn=cint(sValues(i))
		'response.write "combo column id:"&sValues(i)&"-"&fnRST(iIndexColumn)&"</br>"
		sValueDisplayed=sValueDisplayed&fnRST(iIndexColumn)&" "
	next
	s=s&"<option value="""&fnRST(iBoundColumn)&""">"&sValueDisplayed&"</option>"&vbcr
	fnRST.movenext
loop
FnCreateCombo=s&"</select>"&vbcr
end function

function rda(sURL)
response.redirect(sUrl)
end function

function FieldEncoder(idField)
set rs=server.createobject("adodb.recordset")
sSQL=" P_Sel_field_lookup_EncID "&idField&","&idb
rs.open sSQL,sMDB
if rs.eof then
	'do something
end if
do until rs.eof
	sUID=rs(0)
loop
rs.close
set rs=nothing
FieldEncoder=sUID
end function

function PositionField(iFieldID)


end function

'Pass the name of the file to the function.
function CusInputField(idField)
'customer clicks on field div popup pop's up to hello there please input data as specified from field_lookup
'ok so the first thing with this is to create the encrypter.
s=""
'so now create entity with included Jscript for update
sUID=FieldEncoder(idField)
set rstDD=server.CreateObject("ADODB.recordset")
rstDD.open "P_Sel_field_lookup_ByID "&ifldID,sMDB
s=sUID&":"
s=s&Converttohtml_Input(rstDD ("Field_Name"),sVal,rstDD("FIELD_TYPE_DES"),rstDD("required"),rstDD("FIELD_LENGTH"),rstDD ("readonly"),sClass,rstDD("IS_Shown"),rstDD("File_Upload"),strKey)
s=s&" style=""width:"&iPixelWidth&"px"" id="""&sUID&iKey&""" class=""htmlInput"" onkeypress=""ufkd(event)"">"
CusInputField=s
rstDD.close
set rstDD=nothing
end function

Function URLDecode(sConvert)
    Dim aSplit
    Dim sOutput
    Dim I
    If IsNull(sConvert) or sConvert="" Then
       URLDecode = ""
       Exit Function
    End If
	
    ' convert all pluses to spaces
    sOutput = REPLACE(sConvert, "+", " ")
    ' next convert %hexdigits to the character
    aSplit = Split(sOutput, "%")
    If IsArray(aSplit) Then
      sOutput = aSplit(0)
      For I = 0 to UBound(aSplit) - 1
        sOutput = sOutput & _
          Chr("&H" & Left(aSplit(i + 1), 2)) &_
          Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
      Next
    End If
    URLDecode = sOutput
End Function


function Check_Session_Ajax(iUserID)
if len(iUserID)=0 then
	response.write "redirect^/login.asp?msgtxt=ABX300"
	response.end
end if
end function

function itemListView(rs,iItemShow)
'photo
'Could be a variety of possiblities depending on login...
'ideally staff should have direct access to update title price etc.
'Idea chose fields depending on 
'X=ifa(rs.source)
if iItemShow mod 2=0 then
	sClass="itemListClear"
else
	sClass="itemListGray"
end if
s=""
if bSiteBuyNowTM=1 then
	sLinkItem=rs("item_link_TM") 'link to item on TM
else
	sLinkItem=rs("url_link")
end if
s=s&"<tr class="""&sClass&"""><td>"
s=s&"<div class=""itemListImage"">"
s=s&"<a href="""&sLinkItem&""" class=""noBorder"" ><img src=""/images/db"&iDB&"/stock/x-small/"&rs("part_number")&".jpg"" alt=""item photo""/></a>"
s=s&"</div>"
s=s&"</td>"
s=s&"<td>"
s=s&"<div class=""itemListTitleHolder"">"
s=s&"<a href="""&sLinkItem&""">"&rs("title")&"</a>"
if rs("condition")=1 and bShowTue=true then s=s&"<a href="""&sLinkItem&"""><img src=""/images/newItem.gif"" title=""New item"" alt=""New item"" /></a>"
s=s&"</div>"
s=s&"<div class=""itemListPC"">"
s=s&"<span class=""infoText"">"&rs("stock_quantity")&" piece in stock</span></br>"
if rs("is_fav")=0 then
	s=s&"<div class=""Favlist""><a class=""spButton FavlistButton"" id=""itm"&rs("id")&""" href=""/MyPages/AddToFavouritesNoScript.asp?id=itm"&rs("id")&""" onclick=""saveFav(this.id);return false;"" >Save to favourites</a></div>"
else
	s=s&"<div class=""Favlist""><a class=""spButton favSaved"" id=""itm"&rs("id")&""">Favourite</a></div>"
end if
s=s&"</div>"
s=s&"</td>"
s=s&"<td>"
bShowRetail=true
bShowBulk=false
if sScriptName="RECENTLYVIEWED.ASP" then
	bShowRetail=true
	bShowBulk=false
end if
s=s&pricingDiv(rs,bShowRetail,bShowBulk)
s=s&"</td>"
s=s&"<td>"
s=s&"<div class=""itemListBuyOptions"">"
s=s&"	<div class=""spacer""> </div>"
s=s&"	<div class=""QbuyList"" onclick=""location.href='/process_order.asp?id="&rs("id")&"'"" style=""float:right;""><a id=""quickbNButton"" href=""/process_order.asp?id="&rs("id")&""">Add to Cart</a></div>"
if rs("stock_quantity")>0 then 
	s=s&"<div class=""itemListInStock""><span class=""stockQuantityText"">In Stock</span></div>"
else
	s=s&"<div class=""itemListOutStock""><span class=""stockQuantityText"">Sold Out&nbsp;</span><a href=""/help/contactByEmail.asp?sid="&rs("id")&""" class=""iconLink""><img src=""/images/icons/arrow_refresh.png"" onclick="""" ><span class=""smallLink"" style=""margin-left:5px;"">Request restock</span></a></div>"
end if
s=s&"</div>"
s=s&"</td></tr>"
s=s&"<tr><td class=""itemListSeperator"" colspan=""4"">"
s=s&"</td></tr>"
itemListView=s
end function

function pricingDiv(rs,bShowRetail,bPriceShowBulk)
'bPriceShowBulk=true
'bShowRetail=true
if session("TaxMult")="" then session("TaxMult")=1
sp=""
'x=rwe(rs.source)
sp=sp&"<div class=""itemListPriceHolder"">"
sp=sp&"<table style=""margin-left:auto;margin-right:auto;"">"
if bShowRetail=true then
	bRetailFailSave=true
	if bRetailFailSave=true then
		if isnull(rs("retail")) then mRetail=0
		if rs("retail")-rs("price")<0 then
			mRetail=rs("price")+rs("price")*1
		else
			mRetail=rs("retail")
		end if
	end if
	if isnull(mRetail) then mRetail=0
	'x=rwb(mRetail*session("TaxMult"))
	sp=sp&"<tr><td class=""galleryMainPriceRetail""><span class=""noFiveBestPrice"">Retail&nbsp;Price </span></td>"
	sp=sp&"<td class=""galleryMainPrice""><span>$"&formatnumber(mRetail*session("TaxMult"),2)&"</span></td>"
	sp=sp&"</tr>"
	sp=sp&"<tr><td class=""galleryMainPriceYouSave""><span class=""priceYourDiscText"">Your&nbsp;Discount </span></td>"
	iPrice=rs("My_price")
	if isnull(rs("My_price")) then iPrice=0
	yourPrice=((mRetail-cdbl(iPrice))*session("TaxMult"))
	'x=rwb(yourPrice)
	'yourPrice=cdbl(yourPrice)
	sp=sp&"<td class=""galleryMainPrice""><span class=""priceYourDiscValue"">$-"&formatnumber(yourPrice,2)&"</span></td>"
	sp=sp&"</tr>"
end if
sp=sp&"<tr><td class=""galleryMainPriceRetail""><span class=""priceYourPriceText"">Your&nbsp;Price </span></td>"
sp=sp&"<td class=""galleryMainPrice""><span class=""priceYourPriceValue"">$"&formatnumber(cdbl(iPrice)*session("TaxMult"),2)&"</span></td>"
sp=sp&"</tr>"
if bPriceShowBulk=true then
	sp=sp&"<tr class=""galleryPriceHolderSub"">"
	sp=sp&"<td class=""galleryMainPriceDeal1Break""><span class=""noFourBestPrice"">"&rs("Cat_Disc_Break3")&" +</span></td>"
	sp=sp&"<td class=""galleryMainPricedealer3"">$"&formatnumber(rs("dealer3")*session("TaxMult"),2)&"</td>"
	sp=sp&"</tr><tr><td class=""galleryMainPriceDeal2Break""><span class=""noThreeBestPrice"">"&rs("Cat_Disc_Break2")&" +</span></td>"
	sp=sp&"<td class=""galleryMainPricedealer2"">$"&formatnumber(rs("dealer2")*session("TaxMult"),2)&"</td>"
	sp=sp&"</tr><tr><td class=""galleryMainPriceDeal3Break""><span class=""noTwoBestPrice"">"&rs("Cat_Disc_Break1")&" +</span></td>"
	sp=sp&"<td class=""galleryMainPricedealer1"">$"&formatnumber(rs("dealer1")*session("TaxMult"),2)&"</td>"
sp=sp&"</tr>"
end if
sp=sp&"</table></div>"

pricingDiv=sp
end function

function CustomPrices(rs)
'x=rwe(rs.source)
bPriceShowBulk=true
s=""
iItem=rs("id")
if bSiteBuyNowTM=1 then
	sLinkItem=rs("item_link_TM") 'link to item on TM
else
	sLinkItem=rs("url_link")
end if
s=s&"<tr><td>"
s=s&"<div class=""itemListImage"">"
s=s&"<a href="""&sLinkItem&""" class=""noBorder""><img src=""/images/db"&iDB&"/stock/x-small/"&rs("part_number")&".jpg"" alt=""item photo""/></a>"
s=s&"</div>"
s=s&"</td>"
s=s&"<td>"
s=s&"<div class=""itemListTitleHolder"">"
s=s&"<a href="""&sLinkItem&""">"&rs("title")&"</a>"
if rs("condition")=1 and bShowTue=true then s=s&"<a href="""&sLinkItem&"""><img src=""/images/newItem.gif"" title=""New item"" alt=""New item"" /></a></br>"
if len(rs("comment"))>0 then
	s=s&"<div class=""staffNote"" title=""last edited "&rs("date_edited")&""">"&rs("noteDetail")&"</div>"
end if
s=s&"<div id=""favNoteHolder"&iItem&""">"
s=s&ShowNoteText(rs("id"))
s=s&"</div>"
s=s&"<div class=""listEmlOptions"">"
s=s&"<table class=""eml""><tbody><tr>"
s=s&"<td id=""td"&rs("id")&""" class=""fltLeft"">"
s=s&"<img src=""/images/icons/email_attach.png"" onclick=""emlFavShow(this.id,event)"" class=""emlIcon"" id=""ei"&rs("id")&""" alt=""eml"">"
s=s&"</td>"
s=s&"<td><table><tr><td>"
s=s&"<a href=""/MyPages/EmailOptionsNoScript.asp?id="&rs("id")&""" id=""emlLink"&rs("id")&""" class=""emailAjaxIcon nowrap"" onclick=""return:false;"">Email Options</a>"
s=s&"</td><td>"
s=s&"<div class=""listEmlOptions"">"
s=s&"<table class=""notes""><tbody><tr>"
s=s&"<td id=""td"&rs("id")&""" class=""fltLeft"">"
'<img alt="note" id="note167" class="noteIcon" onclick="addInfo('fav',this.id,event)" src="/images/icons/note_add.png">
s=s&"<img src=""/images/icons/note_add.png"" onclick=""addInfo('fav',this.id,event)"" class=""noteIcon"" id=""note"&rs("id")&""" alt=""note"">"
s=s&"</td>"
s=s&"<td><a href=""/MyPages/AddNoteOptionsNoScript.asp?id="&rs("id")&""" id=""emlLink"&rs("id")&""" class=""emailAjaxIcon"" onclick=""return:false;"">Add Notes</a></td>"
s=s&"</tr>"
s=s&"</tbody></table>"
s=s&"</div>"
s=s&"</td>"
s=s&"</tr>"
s=s&"<tr>"
s=s&"<td colspan=""5"" style=""width:100px"">"
s=s&"<span class=""infoText"" id=""emlOptei"&rs("id")&""">set to ("&rs("eml_text")&")</span>"
s=s&"</td></tr>"
s=s&"</tbody></table>"
s=s&"</div>"
s=s&"</div>"
s=s&"</td></tr></table>"
s=s&"<td>"
s=s&"<div class=""itemListPriceHolder""><table style=""margin-left:auto;margin-right:auto;"">"
s=s&"<tr>"
s=s&"<td class=""galleryMainPrice"">"&pricingDiv(rs,true,session("show_bulk_prices"))&"</td>"
s=s&"</tr>"
s=s&"</table></div>"
s=s&"</td>"
s=s&"<td>"
s=s&"<div class=""QbuyList"" onclick=""location='/process_order.asp?id="&rs("id")&"'"">"
s=s&"<div class=""spacer""> </div>"
s=s&"<div class=""""><a class="""" href=""/process_order.asp?id="&rs("id")&""">Add to Cart</a></div>"
s=s&"</div>"
s=s&"</td></tr>"
s=s&"<tr><td class=""itemListSeperator"" colspan=""4"">"
s=s&"</td></tr>"

CustomPrices=s
end function






function rpad(sNumber,iLentoPadTo)
for i=1 to iLentoPadTo-len(sNumber)
	rpad=rpad&"0"
Next
rpad=rpad&sNumber
end function  

function itemGalleryView(rs,iItemShow)
'x=rwe(rs.source)
s=""
if bSiteBuyNowTM=1 then
	sLinkItem=rs("item_link_TM") 'link to item on TM
else
    'sTtitleLink="/"&replace(rs("title")," ","-")&"-"&sSiteAbrivation&rpad(rs("id"),6)
    'sTtitleLink=replace(sTtitleLink,"--","-")
    'sTtitleLink=replace(sTtitleLink,"-,-",",")
   ' sTtitleLink=replace(sTtitleLink,",-","-")
    'sTtitleLink=replace(sTtitleLink,"-,-",",")
	'sLinkItem=replace(rs("urllink"),"/category-"&rs("cat_uid"),sTtitleLink)
	sLinkItem=rs("url_link")
	'x=rwe(sLinkItem)
end if
if iItemShow mod 3=0 then
	sClass="itemGalleryRight"
else
	sClass="itemGallery"
end if
s=s&"<div class="""&sClass&""">"
s=s&"<div class=""itemGalleryImage"">"
if rs("has_photo") then
	sPhoto=rs("part_number")
else
	sPhoto="no-photo"
end if
s=s&"<a href="""&sLinkItem&""" class=""imageLinks""><img class=""noBorder"" src=""/images/db"&iDB&"/stock/small/"&sPhoto&".jpg"" alt=""item Photo"" width=""200"" height=""150""/></a>"
s=s&"</div>"
s=s&"<div class=""galleryTitleHolder"">"
s=s&"<a href="""&sLinkItem&""">"&rs("title")&"</a>"
if rs("condition")=1 and bShowTue=true then s=s&"<a href="""&sLinkItem&"""><img src=""/images/newItem.gif"" title=""New item"" alt=""New item"" /></a>"
s=s&"</div>"
s=s&"<div class=""galleryPriceHolder""><table style=""margin-left:auto;margin-right:auto;"">"
s=s&"<tr><td class=""galleryMainPriceHead"">"&pricingDiv(rs,true,session("show_bulk_prices"))&"</td>"
s=s&"</tr>"
s=s&"</table></div>"
if rs("stock_quantity")>0 then 
	s=s&"<div class=""itemGalleryInStock""><span class=""stockQuantityText"">In Stock</span></div>"
else
	s=s&"<div class=""itemGalleryOutStock""><span class=""stockQuantityText"">Sold Out&nbsp;</span><a href=""/help/contactByEmail.asp?sid="&rs("id")&""" class=""iconLink""><img src=""/images/icons/arrow_refresh.png"" onclick="""" ><span class=""smallLink"" style=""margin-left:5px;"">Request restock</span></a></div>"
end if
s=s&"<div class=""galleryBuyOptions"">"
if rs("is_fav")=0 then
	s=s&"<div class=""Favlist""><a class=""spButton FavlistButton"" id=""itm"&rs("id")&""" href=""/MyPages/AddToFavouritesNoScript.asp?id=itm"&rs("id")&""" onclick=""saveFav(this.id);return false;"" >Save to favourites</a></div>"
else
	s=s&"<div class=""Favlist""><a class=""spButton favSaved"" id=""itm"&rs("id")&""">Favourite</a></div>"
end if
s=s&"<div class=""spacer"">&nbsp;</div>"
s=s&"<div id=""quickbNButton"" onclick=""location.href='/process_order.asp?id="&rs("id")&"'"" style=""float:right;""><a id=""quickbNButton"" href=""/process_order.asp?id="&rs("id")&""">Add to Cart</a></div>"
s=s&"</div>"
s=s&"</div>"
itemGalleryView=s
end function



function itemMarketView(rs,iItemShow)
'x=rwe(rs.source)
s=""
if bSiteBuyNowTM=1 then
	sLinkItem=rs("item_link_TM") 'link to item on TM
else
    'sTtitleLink="/"&replace(rs("title")," ","-")&"-"&sSiteAbrivation&rpad(rs("id"),6)
    'sTtitleLink=replace(sTtitleLink,"--","-")
    'sTtitleLink=replace(sTtitleLink,"-,-",",")
   ' sTtitleLink=replace(sTtitleLink,",-","-")
    'sTtitleLink=replace(sTtitleLink,"-,-",",")
	'sLinkItem=replace(rs("urllink"),"/category-"&rs("cat_uid"),sTtitleLink)
	sLinkItem=rs("url_link")
	'x=rwe(sLinkItem)
end if
if iItemShow mod 3=0 then
	sClass="itemMarketingGalleryRight"
else
	sClass="itemMarketingGallery"
end if
s=s&"<div class="""&sClass&""">"
s=s&"<div class=""itemGalleryImage"">"
if rs("has_photo") then
	sPhoto=rs("part_number")
else
	sPhoto="no-photo"
end if
s=s&"<a href="""&sLinkItem&""" class=""imageLinks""><img class=""noBorder"" src=""/images/db"&iDB&"/stock/small/"&sPhoto&".jpg"" alt=""item Photo"" width=""200"" height=""150""/></a>"
s=s&"</div>"
s=s&"<div class=""galleryTitleHolder"">"
s=s&"<a href="""&sLinkItem&""">"&rs("title")&"</a>"
if rs("condition")=1 and bShowTue=true then s=s&"<a href="""&sLinkItem&"""><img src=""/images/newItem.gif"" title=""New item"" alt=""New item"" /></a>"
s=s&"</div>"
s=s&"<div class=""galleryPriceHolder""><table style=""margin-left:auto;margin-right:auto;"">"
s=s&"<tr><td class=""galleryMainPriceHead"">"&pricingDiv(rs,true,session("show_bulk_prices"))&"</td>"
s=s&"</tr>"
s=s&"</table></div>"
s=s&"</div>"
itemMarketView=s
end function


function showNoteText(iStockID)
set rsNote=server.Createobject("adodb.recordset")
'x=rwe("P_Sel_customerNotesBycusSys "&session("cus_sys")&","&iStockID)
rsNote.open "P_Sel_customerNotesBycusSys "&session("cus_sys")&","&iStockID,sMDB
	sn=""
do until rsNote.eof
	sn=sn&"<div class=""favNotes"" id=""dvNote"&rsNote("id")&""">"
	sn=sn&"<table style=""width:100%""><tbody><tr>"
	sn=sn&"<td id=""note"&rsNote("id")&""" class=""fltLeft"">"
	sn=sn&"<div id=""noteTxt"&rsNote("id")&""">"&replace(rsNote("note_text"),chr(10),"</br>")&"</div></br>"
	sn=sn&"<span class=""subText"">note added on "&DateFriendly(rsNote("date_inserted"),1)&".</span>"
	sn=sn&"</td><td class=""favNoteActions"">"
	sn=sn&"<img src=""/images/icons/note_edit.png"" style=""float:left;"" onclick=""editNotes('fav','noteTxt"&rsNote("id")&"',"&rsNote("related_key")&",event)"">"
	sn=sn&"<img src=""/images/icons/note_delete.png"" style=""float:left;"" onclick=""deleteNote("&rsNote("id")&")"">"
	sn=sn&"</td>"
	'sn=sn&"<td><a href=""/MyPages/AddNoteOptionsNoScript.asp?id="&rs("id")&""" id=""emlLink"&rs("id")&""" class=""emailAjaxIcon"" onclick=""return:false;"">Add Notes</a></td>"
	sn=sn&"</tr>"
	sn=sn&"</tbody></table>"
	sn=sn&"</div>"
	rsNote.MoveNext
loop
rsNote.close
set rsNote=nothing
showNoteText=sn
'sn=""
end function

function showNoteTextByRs(rsNote)
sn=""
do until rsNote.eof
	sn=sn&"<div class=""favNotes"" id=""dvNote"&rsNote("id")&""">"
	sn=sn&"<table style=""width:100%""><tbody><tr>"
	sn=sn&"<td id=""note"&rsNote("id")&""" class=""fltLeft"">"
	sn=sn&"<div id=""noteTxt"&rsNote("id")&""">"&replace(rsNote("note_text"),chr(10),"</br>")&"</div></br>"
	sn=sn&"<span class=""subText"">note added on "&DateFriendly(rsNote("date_inserted"),1)&".</span>"
	sn=sn&"</td><td class=""favNoteActions"">"
	sn=sn&"<img src=""/images/icons/note_edit.png"" style=""float:left;"" onclick=""editNotes('fav','noteTxt"&rsNote("id")&"',"&rsNote("related_key")&",event)"">"
	sn=sn&"<img src=""/images/icons/note_delete.png"" style=""float:left;"" onclick=""deleteNote("&rsNote("id")&")"">"
	sn=sn&"</td>"
	'sn=sn&"<td><a href=""/MyPages/AddNoteOptionsNoScript.asp?id="&rs("id")&""" id=""emlLink"&rs("id")&""" class=""emailAjaxIcon"" onclick=""return:false;"">Add Notes</a></td>"
	sn=sn&"</tr>"
	sn=sn&"</tbody></table>"
	sn=sn&"</div>"
	rsNote.MoveNext
loop
rsNote.close
set rsNote=nothing
showNoteTextByRs=sn
'sn=""
end function


function itemFavourites(rs,sClass)
bPriceShowBulk=true
s=""
iItem=rs("id")
if bSiteBuyNowTM=1 then
	sLinkItem=rs("item_link_TM") 'link to item on TM
else
	sLinkItem=replace(rs("urllink"),"/category-"&rs("cat_uid"),"/"&CleanCatText(replace(rs("title")," ","-"))&"-"&sSiteAbrivation&rpad(rs("id"),6))
end if
s=s&"<tr class="""&sClass&""">"
s=s&"<td valign=""top""><input type=""checkbox"" class=""chkboxAction selectable"" id=""fvChk"&rs("id")&"""></td>"
s=s&"<td>"
s=s&"<div class=""itemListImage"">"
s=s&"<a href="""&sLinkItem&""" class=""noBorder""><img  src=""/images/db"&iDB&"/stock/x-small/"&rs("part_number")&".jpg"" alt=""item photo""/></a>"
s=s&"</div>"
s=s&"</td>"
s=s&"<td>"
s=s&"<div class=""itemListTitleHolder"">"
s=s&"<span style=""float:left;""><a href="""&sLinkItem&""">"&rs("title")&"</a></span>"
if rs("condition")=1 and bShowTue=true then s=s&"<a href="""&sLinkItem&"""><img src=""/images/common/newItem.gif"" title=""New item"" alt=""New item"" /></a></br>"

s=s&"<div id=""favNoteHolder"&iItem&""">"
s=s&ShowNoteText(rs("id"))
s=s&"</div>"
s=s&"<div class=""listEmlOptions"">"
s=s&"<table class=""eml""><tbody><tr>"
s=s&"<td id=""td"&rs("id")&""" class=""fltLeft"">"
'x=rwe(rs.source)
s=s&"<img src=""/images/icons/note_add.png"" onclick=""addInfo('fav',this.id,event)"" class=""noteIcon"" id=""note"&rs("id")&""" alt=""note"">"
s=s&"</td>"
s=s&"<td><a href=""/MyPages/AddNoteOptionsNoScript.asp?id="&rs("id")&""" id=""emlLink"&rs("id")&""" class=""emailAjaxIcon"" onclick=""return:false;"">Add Notes</a></td>"
s=s&"</td>"
s=s&"</tr>"
s=s&"<tr>"
s=s&"<td id=""td"&rs("id")&""" class=""fltLeft"">"
s=s&"<img src=""/images/icons/email_attach.png"" onclick=""emlFavShow(this.id,event)"" class=""emlIcon"" id=""ei"&rs("id")&""" alt=""eml"">"
s=s&"</td>"
s=s&"<td>"
s=s&"<a href=""/MyPages/EmailOptionsNoScript.asp?id="&rs("id")&""" id=""emlLink"&rs("id")&""" class=""emailAjaxIcon"" onclick=""return:false;"">Email Options</a>"
s=s&"</td>"
s=s&"</tr>"
s=s&"</tbody></table>"
s=s&"</div>"
s=s&"<div class=""indent20 clear""><span class=""infoText"" id=""emlOptei"&rs("id")&""">set to ("&rs("eml_text")&")</span></div>"
s=s&"</div>"
s=s&"</td>"
s=s&"<td style=""vertical-align:top""><div>"
s=s&pricingDiv(rs,true,session("show_bulk_prices"))
's=s&"</td>"
's=s&"<td>"
s=s&"<div class=""itemListBuyOptions"">"
s=s&"<div class=""spacer""> </div>"
s=s&"<table><tr><td>"
if rs("stock_quantity")>0 then
s=s&"<span class=""stockQuantityText"">Availible stock </span><span class=""stockQuantityUnits"">"&rs("stock_quantity")&"</span>"
else
	s=s&"<div class=""itemFavOutStock""><span class=""stockQuantityText"">Sold Out&nbsp;</span></br><a href=""/help/contactByEmail.asp?sid="&rs("id")&""" class=""iconLink""><img src=""/images/icons/arrow_refresh.png"" onclick="""" ><span class=""smallLink"" style=""margin-left:5px;"">Request restock</span></a></div>"
end  if
s=s&"</td><td>"
s=s&"<div id=""quickbNButton"" onclick=""location.href='/process_order.asp?id="&rs("id")&"'"" style=""float:right;"">"
s=s&"<a id=""quickbNButton"" href=""/process_order.asp?id="&rs("id")&""">Add to Cart</a>"
s=s&"</div></td></tr></table>"
s=s&"</div></div>"
s=s&"</td></tr>"
s=s&"<tr><td class=""itemListSeperator"" colspan=""4"">"
s=s&"</td></tr>"
itemFavourites=s
end function

function rpad(sNumber,iLentoPadTo)
for i=1 to iLentoPadTo-len(sNumber)
	rpad=rpad&"0"
Next
rpad=rpad&sNumber
end function  


FUNCTION MXLookupAJAX(host)
  MXLookupAJAX = ""
  Dim objXMLHTTP,strResult
  DIM lines, accepted : accepted = -1
  Set objXMLHTTP = Server.CreateObject("Microsoft.XMLHTTP")
  objXMLHTTP.Open "Get", "/admin/checkEmalValid.asp?domainname=" & host & "&Submit=Submit&t_mx=1", False
  objXMLHTTP.Send
  strResult = objXMLHTTP.ResponseText
  lines = Split(str,chr(13))
  FOR i = 0 TO UBound(lines)
    IF InStr(lines(i),"MX preference = ") > 0 THEN
     accepted = i
     EXIT FOR
    END IF
  NEXT
  IF accepted > -1 THEN
    MXLookupAJAX= "Accepted: " & lines(accepted)
  ELSE
   MXLookupAJAX= "Denied! "
  END IF
END FUNCTION

FUNCTION MXLookup(host)
  SET objShell = Server.CreateObject("Wscript.Shell")
  DIM objExec, strResult
  'SET objExec = objShell.Exec("%comspec% /c nslookup -type=MX " & host)
  'Set objCmd = objShell.Exec("ping " & "www."&host) 
  Set objCmd = objShell.Exec("%comspec% /c nslookup -type=MX " & host) 
  strPResult = objCmd.StdOut.Readall() 
 	MXLookup=strPResult
  'x=rwe(strPResult)
  'strResult = objExec.StdOut.ReadAll
END FUNCTION

Function Add_Mail_User()
Dim objShell
Dim vbsFile

vbsFile = Server.MapPath("\admin\mail\AddUser.vbs")
Set objShell = Server.CreateObject("Wscript.Shell")

objShell.Run vbsFile

Set objShell = Nothing

Response.Write "Done"

end function




Function createColor(iNum)
if isnull(iNum) then 
	createColor="#FF0000"
	exit function
end if
If iNum <1 Then createColor="#FF0000"
If iNum >1 Then createColor="#FF9900"
If iNum >5 Then createColor="#FFAA33"
If iNum >10 Then createColor="#669900"
If iNum >20 Then createColor="#003399"
End Function

Function CeateTxtFile(sFile,sContents)
dim fs,tfile
set fs=Server.CreateObject("Scripting.FileSystemObject")
'x=rwe(sFile)
set tfile=fs.CreateTextFile(sFile,true,true)

tfile.WriteLine(sContents)
tfile.close
set tfile=nothing
set fs=nothing
End Function

Function getFileContents(strIncludeFile)
  Dim objFSO
  Dim objText
  Dim strPage
  'Instantiate the FileSystemObject Object.
  Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
  'Open the file and pass it to a TextStream Object (objText). The
  '"MapPath" function of the Server Object is used to get the
  'physical path for the file.
  Set objText = objFSO.OpenTextFile(Server.MapPath(strIncludeFile))
  'Read and return the contents of the file as a string.
  getFileContents = objText.ReadAll
  objText.Close
  Set objText = Nothing
  Set objFSO = Nothing
End Function

Sub AddNewFolder(sPath, folderName)
	'response.write  "</br>new folder is " &sPath&""&folderName&"</br>"
	'on error resume next
  Dim fso, f, fc, nf
	'on error resume next
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f = fso.GetFolder(sPath)
  if err.number<>0 then response.write  "</br>Error createing:"&folderName
  Set fc = f.SubFolders
  If folderName <> "" then 
    Set nf = fc.Add(folderName)
  Else
    Set nf = fc.Add("New Folder")
  End If
  'on error goto 0
End Sub

function DVWrap(sText,sDVName,sDVClassName)
sReturn="<div id="""&sDVName&""" class="""&sDVClassName&""">"&sText&"</div>"
end function

function CleanCatText(sCat)
if isnull(sCat) then 
	CleanCatText=""
	exit function
end if  
'^(ht|f)tp(s?)\:\/\/[0-9a-zA-Z]([-.\w]*[0-9a-zA-Z])*(:(0-9)*)*(\/?)([a-zA-Z0-9\-\.\?\,\'\/\\\+&amp;%\$#_]*)?$
'regular expression for validating URLS from microsoft
 '^([0-9a-zA-Z]([-.\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\w]*[0-9a-zA-Z]\.)+[a-zA-Z]{2,9})$
'Validates an e-mail address.
sCatPageLink=replace(sCat," \ ","-")
sCatPageLink=replace(sCat,"  "," ")
sCatPageLink=replace(sCat,"   "," ")
sCatPageLink=replace(sCatPageLink," / ","-")
sCatPageLink=replace(sCatPageLink,"\","-")
sCatPageLink=replace(sCatPageLink,"/","-")
sCatPageLink=Replace(sCatPageLink," & ","-")
sCatPageLink=Replace(sCatPageLink,"&","-")
sCatPageLink=replace(sCatPageLink," ","-")
sCatPageLink=replace(sCatPageLink,"""","")
sCatPageLink=replace(sCatPageLink,"'","")
sCatPageLink=replace(sCatPageLink,"_","-")
sCatPageLink=replace(sCatPageLink,"--","-")
sCatPageLink=replace(sCatPageLink,".","")
if right(sCatPageLink,1)="-" and len(sCatPageLink)>1 then sCatPageLink=left(sCatPageLink,len(sCatPageLink)-2)
CleanCatText=sCatPageLink
'x=rwb(CleanCatText)
end function

function MakeColor(strColorType,intRow)
Select Case strColorType
Case "Groups"
	iNum=intRow mod 2
	if iNum=1 then MakeColor="#6495ED"
	if iNum=0 then MakeColor="#FFCC33"
case "random" 
     'Define Colors
     intRandomColor=rnd*16581375
     strHexTemp=hex(intRandomColor)
     MakeColor ="#" & strHexTemp
end select
end function

function strTime(iMinute)
	iHours=formatNumber(iMinute/60,1)
	strTime=iHours 
end function

function Price_Calculator(iCostPrice,iSellPrice,iClientPercent,st_special)
	'used on products.asp for calculating prices
	iLowPriceCeiling=1.02
	if isnull(iClientPercent) or iClientPercent="" or iClientPercent=0 then iClientPercent=1
	if st_special=true then iClientPercent=1-((1-iClientPercent)/2)
	if iSellPrice*iClientPercent<iCostprice*iLowPriceCeiling then 
		iPrice=iCostprice*iLowPriceCeiling
	else
		iPrice=iSellPrice*iClientPercent
	end if
	Price_Calculator=formatNumber(iPrice,2)
end function

Function dbDate(dt) 
		sTime=right("0" & Hour(dt), 2)&":"&right("0" & Minute(dt), 2)&":"&right("0" & Second(dt), 2)
    dbDate = year(dt) & right("0" & month(dt), 2) & right("0" & day(dt),2) & " " & sTime
End Function 

Function Power_Round(iNum,iNearist)
'Rounds number to the nearest multiple as specified
Power_Round=round(iNum/iNearist,0)*iNearist
end Function 

FUNCTION GetNumbers(sText)
iLen=len(sText)
sNum=""
for i=1 to iLen
	sChar=mid(sText,i,1)
	if asc(sChar)>47 and asc(sChar)<58 or sChar="." then
		sNum=sNum&sChar
	end if
next
if len(sNum)>1 then 
	GetNumbers=sNum
else
	GetNumbers="0"
end if
if right(GetNumbers,1)="." then GetNumbers=left(GetNumbers,len(GetNumbers)-1)
end FUNCTION

Function stripInvalid(sVar)
if sVar="" then 
	stripInvalid=""
	exit  function
end if
sVar=replace(sVar,";","")
sVar=replace(sVar,"'","''")
sVar=replace(sVar,"%","")
sVar=replace(sVar,"sys.objects","")
sVar=replace(sVar,"DROP ","")
sVar=replace(sVar,"SELECT ","")
sVar=replace(sVar,"EXEC ","")
 stripInvalid=sVar
end function

Function strip(sVar)
if sVar="" then 
	stripInvalid=""
	exit  function
end if
sVar=replace(sVar,";","")
sVar=replace(sVar,"'","''")
sVar=replace(sVar,"%","")
sVar=replace(sVar,"sys.objects","")
sVar=replace(sVar,"DROP ","")
sVar=replace(sVar,"SELECT ","")
sVar=replace(sVar,"EXEC ","")
 strip=sVar
end function

function FormDataDump(bolShowOutput, bolEndPageExecution)
  s=""
  Dim sItem
  Dim strLineBreak
  If bolShowOutput then
    strLineBreak = "</br>"
  Else
    strLineBreak = vbCrLf
    s=s&"<!--" & strLineBreak
  End If
 ' s=s&"</br>form collection" & strLineBreak
  For Each sItem In Request.Form
    s=s&sItem
    s=s&" - [" & Request.Form(sItem) & "]" & strLineBreak
  Next
  'Display the Request.QueryString collection
  s=s&strLineBreak & strLineBreak
  's=s&"querystring collection" & strLineBreak
  For Each sItem In Request.QueryString
    s=s&sItem
    s=s&" - [" & Request.QueryString(sItem) & "]" & strLineBreak
  Next
  'If we are wanting to hide the output, display the closing
  'HTML comment tag
  If Not bolShowOutput then FormDataDump=strLineBreak & "-->"
  'End page execution if needed
  If bolEndPageExecution then Response.End
  FormDataDump=s
End function

function GetSiteEmails(iDB)
s="<div style=""float:left""><table><tr><td colspan=1 align=left>From:</td><td><select id=""emlFrom"" name=""emlFrom"" style=""width:150px"">"
s=s&"<option value="""&session("staff_email")&""">"&session("staff_email")&"</option>"
if sSiteEmail1<>"" then
s=s&"<option value="""&sSiteEmail1&""">"&sSiteEmail1&"</option>"
end if
if sSiteEmailSales<>"" then
s=s&"<option value="""&sSiteEmailSales&""">"&sSiteEmailSales&"</option>"
end if
s=s&"</select><td></tr></table></div>"
GetSiteEmails=s
'x=rwe(s)
end function

function GetAdminEmails(iDB)
s="<table><tr><td colspan=1 align=left>From:</td><td><select id=""emlFrom"" name=""emlFrom"">"
s=s&"<option value="""&session("staff_email")&""">"&session("staff_email")&"</option>"
if sSiteEmail1<>"" then
s=s&"<option value="""&sSiteEmail1&""">"&sSiteEmail1&"</option>"
end if
if sSiteEmailSales<>"" then
s=s&"<option value="""&sSiteEmailSales&""">"&sSiteEmailSales&"</option>"
end if
s=s&"</select><td></tr></table>"
GetSiteEmails=s
'x=rwe(s)
end function

function ShowChrs(sText)
for x=1 to len(sText)
	sChar=mid(sText,x,1)
	ShowChrs=ShowChrs&"chr("&asc(sChar)&"):"
next
end function

function FindNextSperator(sText)
dim sCommands(20)
'FindNextSperator=sText
FindNextSperator=""
sCommands(0)=" AND "
sCommands(1)=" OR "
sCommands(2)=" ORDER BY "
sCommands(3)=" GROUP BY "
iIndex=10000
for i=0 to 3
	if instr(sText,sCommands(i))<iIndex and instr(sText,sCommands(i))>0 then
		iIndex=instr(sText,sCommands(i))
		FindNextSperator=sCommands(i)
	end if
next
end function


function StartFromText(sFullString,sFindString,bEndOfText)
iMoveTo=instr(sFullString,sFindString)
if iMoveTo=0 or len(sFindString)>=len(sFullString) then
	StartFromText=sFullString
else
	if bEndOfText=1 then iSub=len(sFindString) else iSub=-1
	StartFromText=right(sFullString,len(sFullString)-iMoveTo-iSub)
end if
end function

function isReallyNumeric(str) 
isReallyNumeric = true 
for i = 1 to len(str) 
    d = mid(str, i, 1) 
    if asc(d) < 48 OR asc(d) > 57 then 
        isReallyNumeric = false 
        exit for 
    end if 
next 
end function 
    


function CheckFormPosts()
if request.form.count>0 then 
	CheckFormPosts=1
ELSE
	CheckFormPosts=0
end if
end function

function sURLFromForm(sURLCheck)
	'x=ifa("</br>sURLFromForm")
	sOutPut=""
		if instr(sURLCheck,"?")=0 then
			sPrefix="?"
		ELSE
			sPrefix="&"
		end if
	For Each Item In Request.Form
		fieldName = Item
		fieldValue = Request.Form(Item) 
		'x=rwb(request.form.count)
		sUrlTemp=sPrefix&fieldName&"="&fieldValue  
		'x=ifa(sUrlTemp)
		if instr(sURLCheck,sUrlTemp)=0 then
			sOutPut=sOutPut&sUrlTemp
		end if
		sPrefix="&"
	Next
	sURLFromForm=sURLCheck&sOutPut
end function 

function MakePages(sUrl,iCount,iPgFile,iPagesShow,iCurrPage,sHTM)
	bForm=CheckFormPosts()
	'x=rwe(bForm)
	'used for paging a recordset, note only need to know the following varibles
	'used on products, tm_adverts
	'x=rwe(sUrl)
	'hard code to make a pagefile of 5
	'use sHTM to add htm extension to link
	'x=ifa("</br>"&sUrl)
	if bForm=1 then sUrl=sURLFromForm(sURL)
	'x=ifa(sUrl)
	if iCurrPage=0 then iCurrPage=1
	if iCount>iPgFile then
		if isnumeric(iCurrPage) then iCurrPage=cint(iCurrPage)
		if instr(sUrl,"?")=0 then
			sPrefix="?"
		else
			sPrefix="&"
		end if
		iPgs=int(iCount/iPgFile)+1
		if iCurrPage="" then iCurrPage=0
		iPagesPlus=round(iPagesShow/2,0)
		iPagesMinus=-round(iPagesShow/2,0)
		if iPagesMinus+iCurrPage<1 then
			iPagesPlus=iPagesPlus-iPagesMinus
		end if
		if iPagesPlus+iCurrPage>=iPgs then
			iPagesMinus=iPagesMinus-iPagesPlus
		end if
		if iCurrPage<>0 then
			iPgMax=cint(iCurrPage+iPagesPlus)
			iPGMin=cint(iCurrPage+iPagesMinus)
			if iPgMax=>iPgs then iPgMax=iPgs
			if iPgMin<1 then iPgMin=1
		else
			iPgMax=iPgs
			iPgMin=1
			if iPgMax>iPagesShow then iPgMax=iPagesShow-1
		end if
		
		iToResults=iCurrPage*iPgFile
		if iToResults>iCount then iToResults=iCount
		sHeader="<div class=""pageDiv""><ul><li class=""itemResultsHead"">"&iCount&" results, showing "&(iCurrPage-1)*iPgFile+1&" to "&iToResults&"</li>"
		sHeader=sHeader&"<li class=""itemResultsNav"">"
		sUrl=replace(sUrl,"&p="&iCurrPage-1,"")
		sUrl=replace(sUrl,"&p="&iCurrPage,"")
		sUrl=replace(sUrl,".htm","/")
		sUrl=replace(sUrl,"/page-"&iCurrPage,"")
		for i=iPgMin to iPgMax
			sURL=replace(sUrl,"&p="&i,"")
			if i=iPgMin and iCurrPage>1 and iCurrPage-1>0 then
				sPg="<a href="""&sURL&sPrefix&"p="&iCurrPage-1&"""><< previous </a>"
				sPGH="<a href="""&sURL&"page-"&iCurrPage-1&sHTM&"""><< previous </a>"
			end if
			if i=iCurrPage then sPg=sPg&" <a style=""color :red;font-weight :bold;font-size:13px;text-decoration:underline;"" href="""&sURL&sPrefix&"p="&i&"""> "&i&" </a> "
			if i=iCurrPage then sPGH=sPGH&" <a class=""color :red;font-weight :bold;font-size:13px;text-decoration:underline;"" href="""&sURL&"page-"&i&sHTM&"""> "&i&" </a> "
			if i<>iCurrPage then sPg=sPg& " <a href="""&sURL&sPrefix&"p="&i&""">"&i&"</a> "
			if i<>iCurrPage then sPGH=sPGH&" <a href="""&sURL&"page-"&i&sHTM&""">"&i&"</a> "
			if i=iPgMax and iCurrPage+1<=iPgMax then
				sPg=sPg&"<a href="""&sURL&sPrefix&"p="&iCurrPage+1&"""> next>> </a>"
				sPGH=sPGH&"<a href="""&sURL&"page-"&iCurrPage+1&sHTM&"""> next>> </a>"
			end if
		next
		sPGH=sPGH&"</li></ul>"
		sPg=sPg&"</li></ul>"
	ELSE
		sHeader="<div class=""pageDiv"">"&iCount&" results, showing 1 to "&iCount&""
	end if
	sFooter="</div>"
	if ucase(sHTM)<>".ASP" then 
		MakePages=sHeader&sPGH&sFooter&"<div class=""clear""></div>"&"<div class=""spacer20""></div>"
	else
		MakePages=sHeader&sPg&sFooter&"<div class=""clear""></div>"&"<div class=""spacer20""></div>"
	end if
	'x=rwe(MakePages)
end function




function ValidateFields(rsDD,sKey)
'check data if it passes validation then allow save.
dim iErrors
bFieldPass=true
with rsDD
do until.eof
	iFldID=.fields("FL_ID")
	'Check field by Field
	sFieldName=.fields("field_name")
	sValue=request("i"&sKey&"f"&iFldID)
	sValidation=.fields("field_type_des")
	sValText=.fields("Validation_Text")
	iMaxLength=.fields("field_length")
	bReq=.fields("Required")
	sPass=false
 	'x=rwb("i"&sKey&"f"&iFldID&":"&sValue)
 	if rsDD("key_field")=true then 
 		iKeyFieldID=rsDD("fl_ID")
 		if sValue="" then sValue=0
 	end if
	Select case sValidation
		case "bit"
			sPass=true
		case "int"
			if isnumeric(sValue)=true then
				if sValue>-2147483648 and sValue<2147483647 then
					sPass=true					
				else
					sErrCode=sErrCode & "<e"&iFldID&">" & sValText & "</e"&iFldID&">"
					iErrors=iErrors+1
				end if
			else
				if bReq=false and sValue="" then
					sPass=true			
				else
					sErrCode=sErrCode & "<e"&iFldID&">" & sValText & " and is Required</e"&iFldID&">"
					iErrors=iErrors+1
				end if
			end if
		case "smallint"
			if isnumeric(sValue)=true then
				if sValue>-32768 and sValue<32768 then
					sPass=true					
				else
					sErrCode=sErrCode & "<e"&iFldID&">" & sValText & "</e"&iFldID&">"
					iErrors=iErrors+1
				end if
			else
				if bReq=false and sValue="" then
					sPass=true			
				else
					sErrCode=sErrCode & "<e"&iFldID&">" & sValText & " and is Required</e"&iFldID&">"
					iErrors=iErrors+1
				end if
			end if
		case "smallint"
			if isnumeric(sValue)=true then
				if sValue>-1 and sValue<256 then
					sPass=true					
				else
					sErrCode=sErrCode & "<e"&iFldID&">" & sValText & "</e"&iFldID&">"
					iErrors=iErrors+1
				end if
			else
				if bReq=false and sValue="" then
					sPass=true			
				else
					sErrCode=sErrCode & "<e"&iFldID&">" & sValText & " and is Required</e"&iFldID&">"
					iErrors=iErrors+1
				end if
			end if	
		case "date","datetime","smalldatetime"
			if isdate(sValue)=true then
				sPass=true				
			else
				if bReq=false and sValue="" then
					sPass=true
				else
					sErrCode=sErrCode & "<e"&iFldID&">This field must have a date value.</e"&iFldID&">"
					iErrors=iErrors+1
				end if
			end if
		case else 
			if len(sValue)>iMaxLength then
				sPass=false
				sErrCode=sErrCode & "<e"&iFldID&">You have entered to much text.  Maximum allowed is "&iMaxLength&". Please reduce by "&abs(iMaxLength-len(sValue))&" characters.</e"&iFldID&">"
			else
				if bReq=true and (sValue="" or sValue="</br>") then
					sErrCode=sErrCode & "<e"&iFldID&">This field must have a value.</e"&iFldID&">"
					iErrors=iErrors+1
				else
					sPass=true			
				end if		
			end if
	end select

	if sPass=false then bFieldPass=false
	if sPass=false then 
			'response.write "No Pass for :" & sFieldName & "=" & sValue &"</br>"
			'response.end
	end if
	.movenext
loop
rsDD.movefirst 
'response.write sPassVal
'response.end
end with
ValidateFields=sErrCode
if ValidateFields<>"" then 
	ValidateFields="<validate><failed><table class=""tResultsFail""><tr><td><img src=""/styles/images/error.gif""></td><td><span class=""error"" />Please Check the form "&iErrors&" errors were found.</span></td></tr></table></failed>"&ValidateFields&"</validate>"
else 
	ValidateFields=""
end if
end function


sub SetTMOptions(iDB)


set rTMOptions=server.createobject("ADODB.recordset")
rTMOptions.open "P_Sel_TM_Logo_options "&iDB,sMDB
sTMHeader=""
sTMFooter=""
'bApply_logo=false
if not rTMOptions.eof then
	sTMHeader=rTMOptions("TMO_TM_Header")
	sTMFooter=rTMOptions("TMO_TM_Footer")
	bApply_logo=rTMOptions("TMO_Apply_logo") 
	sLogo_Name=rTMOptions("TMO_Logo_Name")
	iLogo_Top=rTMOptions("TM_Logo_Top") 
	iLogo_Left=rTMOptions("TM_logo_Left")
	iLogo_Transparancy=rTMOptions("TM_Logo_Transparancy")
	sLogo_Height=rTMOptions("TM_Logo_Height")
	sLogo_Width=rTMOptions("TM_logo_Width")
	btmo_show_bulk_pricing=rTMOptions("tmo_show_bulk_pricing")
	Out_of_stock_Default_Code=rTMOptions("Out_of_stock_Default_Code")
	iTMNoPhoto=rTMOptions("no_photo_id")
	sLogoOverlayPath=sSitePath &"\Images\DB"&iDB&"\logos\"&sLogo_Name
end if
rTMOptions.close
set rTMOptions=nothing
end sub

function rwSQL(sVar)
response.write "<div class=""myInfoWarning"" style=""margin-left:20px;float:left;""><div class=""infoIcon""></div>"&sVar&"</div>"
end function


function RW(sVar)
response.write sVar
end function
function RWE(sVar)
response.write sVar
response.end
end function

function RWError(sVar)
response.write "<div class=""myInfoWarning"" style=""margin-left:20px;float:left;""><div class=""errorIcon""></div>"&sVar&"</div>"
end function
function RWInfo(sVar)
response.write "<div class=""myInfoMsg""><div class=""infoIcon"">&nbsp;</div>"&sVar&"</div>"
end function

function RWBlog(sVar)
response.write "<div class=""myInfoBlog"">"&sVar&"</div>"
end function
function RWBlogNS(sVar)
response.write "<div style=""padding: 10px 16px;background: #EDF4FF;border: 1px solid #CDD4EF;font-size: 13px;line-height: 18px;color: #222;"">"&sVar&"</div>"
end function
function RWB(sVar)
response.write "</br>"&sVar&"</br>"
end function
function RWBE(sVar)
response.write "</br>"&sVar&"</br>"
response.end
end function

function RWD(sVar)
if session("debug")=true then
response.write "<div class=""myInfoWarning"" style=""margin-left:20px;float:left;""><div class=""infoIcon""></div>"&sVar&"</div>"
end if
end function
function RWDE(sVar)
if session("debug")=true then
response.write "<div class=""myInfoWarning"" style=""margin-left:20px;float:left;""><div class=""infoIcon""></div>"&sVar&"</div>"
	response.end
end if
end function

Function ifa(sText)
if session("staff_sec_level")>9 and session("debug")=true then
	'x=rwe(session("debug"))
	response.write "<div class=""myInfoWarning"" style=""margin-left:20px;float:left;""><div class=""infoIcon""></div>"&sText&"</div>"
end if
end function

Function Ifae(sText)
if session("staff_sec_level")>9 and session("debug")=true then
	response.write "<div class=""myInfoWarning"" style=""margin-left:20px;float:left;""><div class=""infoIcon""></div>"&sText&"</div>"
	response.end
end if
end function

function SelectLast12Months
set rsDate=server.createobject("adodb.recordset")
sSelect=""
sSQL="P_SEL_Last12Months "
'x=rwe(sSQL)
rsDate.open sSQL,sMDB
if not rsDate.eof then
	sSelect="<select name=""bDate"" id=""bDate"" onchange="""">"
end if
dPrimaryDate="01-"&rsDate(0)&"-"&rsDate(1)
do until rsDate.eof
	sSelect=sSelect&"<option value="""&rsDate(0)&"-"&rsDate(1)&""""
	if request("bDate")=rsDate(0)&"-"&rsDate(1) then sSelect=sSelect&" Selected"
	sSelect=sSelect&">"&rsDate(0)&"-"&rsDate(1)&"</option>"
	rsDate.movenext
	'x=rwe(sSelect)
loop
sSelect=sSelect&"</select>"	
rsDate.close
set rsDate=nothing
if sSelect="</select>" then sSelect=""
SelectLast12Months=sSelect
end function

function SelectLast12MonthsXero
set rsDate=server.createobject("adodb.recordset")
sSelect=""
sSQL="P_SEL_Last12MonthsTax "
'x=rwe(sSQL)
rsDate.open sSQL,sMDB
if not rsDate.eof then
	sSelect="<select name=""bDate"" id=""bDate"" onchange="""">"
end if
dPrimaryDate="01-"&rsDate(0)&"-"&rsDate(1)
do until rsDate.eof
	sSelect=sSelect&"<option value=""01-"&rsDate(0)&"-"&rsDate(1)&""""
	if request("bDate")="01-"&rsDate(0)&"-"&rsDate(1) then sSelect=sSelect&" Selected"
	sSelect=sSelect&">01-"&rsDate(0)&"-"&rsDate(1)&"</option>"
	rsDate.movenext
	'x=rwe(sSelect)
loop
sSelect=sSelect&"</select>"	
rsDate.close
set rsDate=nothing
if sSelect="</select>" then sSelect=""
SelectLast12MonthsXero=sSelect
end function

'***********************************************************************************************
function sSQL_Field_converter(sFieldType,sValue)
Select case sFieldType
	case "nvarchar","nchar","varchar","char"
		strFieldValue=replace(sValue,"'","''")
		strFieldValue=replace(strFieldValue,"^^","+")
		strFieldSeperator="'"
		case "datetime","date","smalldatetime","datetime2","datetimeoffset"
		if len(sValue)=0 then
			strFieldValue="Null"
			strFieldSeperator=""
		else
			if sFieldType="date" then
				intMaxFieldLength=11
				intFieldLength=15
				sDate=CvbShortdateTime(sValue,false)
			end if
			if sFieldType="datetime" then
				intMaxFieldLength=24
				intFieldLength=28
				sDate=CvbShortdateTime(sValue,true)
			end if	
			if sFieldType="smalldatetime" then
				intMaxFieldLength=19
				intFieldLength=23
				sDate=CvbShortdateTime(sValue,true)
			end if			
			if sFieldType="datetime2" then
				intMaxFieldLength=28
				intFieldLength=32
				sDate=CvbShortdateTime(sValue,true)
			end if	
			if sFieldType="datetime2" then
				intMaxFieldLength=35
				intFieldLength=39
				sDate=CvbShortdateTime(sValue,true)
			end if
			strFieldValue=sDate
			strFieldSeperator="'"
		end if
	case "bit"
		if ucase(sValue)="TRUE" then strFieldValue="1" else strFieldValue="0"
		'response.write ":" & sValue &""
		'response.end
	case else
		if sValue="" then sValue=0
		strFieldSeperator="'"
		strFieldValue=replace(sValue,"$","")
End Select
sSQL_Field_converter=strFieldSeperator & strFieldValue & strFieldSeperator
end function
'***********************************************************************************************
function DateFriendly(strDate,bTime)
if isdate(strDate) then
	DateFriendly=right("0"&day(strDate),2)&"-" & MonthName(month(strDate),true) & "-" & year(strDate)
	sAMPM="am, "
	iHour=hour(strDate)
	if iHour>11 then 
		iHour=iHour-12
		if iHour=0 then iHour=12
		sAMPM="pm, "
	end if
	if bTime then DateFriendly=" "&iHour&":"&right("0"&minute(strDate),2)&" "&sAMPM& DateFriendly
else
	DateFriendly=strDate
end if
end function

function CvbShortdateTime(strDate,bTime)
if isdate(strDate) then
	CvbShortdateTime=year(strDate)& " " & MonthName(month(strDate),true) & " " & right("0"&day(strDate),2)
	if bTime then CvbShortdateTime=CvbShortdateTime&" "&right("0"&hour(strDate),2)&":"&right("0"&minute(strDate),2)&":"&second(strDate)
else
	CvbShortdateTime=strDate
end if
end function
'***********************************************************************************************
function CvbShortdate(strDate)
if isdate(strDate) then
	CvbShortdate=day(strDate) & "-" & MonthName(month(strDate),true) & "-" & right(year(strDate),2)
else
	CvbShortdate=strDate
end if
end function
'***********************************************************************************************
function TMShortDate(strDate)
if isdate(strDate) then
	CvbShortdate=right("0"&day(strDate),2) & "/" & right("0"&month(strDate),2) & "/" & right(year(strDate),4)
else
	CvbShortdate=strDate
end if
end function
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'-----------------------------------------------Capitalize First Letters----------------------------''
function capitalize_first_letters(strText)
if len(strText)=0 or isnull(strText) then exit function
a=1
strText=lcase(trim(strText))
if right(strText,1)="," then strText=left(strText,len(strText)-1)
strText=ucase(left(strText,1)) & right(strText,len(strText)-1)
on error resume next
do until instr(a,strText," ")=0
	b=instr(a,strText," ")
	strText=left(strText,b) & ucase(mid(strText,b+1,1)) & right (strText,len(strText)-b-1)
	a=b+1
loop
a=1
do until instr(a,strText,",")=0
	b=instr(a,strText,",")
	strText=left(strText,b) & ucase(mid(strText,b+1,1)) & right (strText,len(strText)-b-1)
	a=b+1
loop
if Err.Number<>0 then 
	response.write(strText)
	response.end
end if
on error goto 0
capitalize_first_letters=strText
end function
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'-------------------------------------------Capitalize First Letter only--------------------------------------------''
function capitalize_first_letter(strText)
if len(strText)=0 or isnull(strText) then exit function
strText=lcase(trim(strText))
strText=ucase(left(strText,1)) & right(strText,len(strText)-1)
capitalize_first_letter=strText
end function
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
Function Gender_Decode(strGender)
Select case strGender
	case "U"
		 Gender_Decode="Unknown"
	case "M"
		Gender_Decode="Male"
	case "F"
		Gender_Decode="Female"
end select
End Function
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
Function DisplaySQL(strSQL,bInline,iDecimal_Places,bHighLight)
dim rsTemp
set rsTemp=server.createobject("ADODB.recordset")

DisplaySQL=""
sFont=""
sFontClose=""
sRowHeaderClass=""
sLink=""
if bInline=0 then sHTML="<table class='high_light_white' width=""300px"">"
if bInline=0 then sRowHeaderClass="high_light_header"
if bHighLight=1 then
	 sFont="<font color=red><b>"
	 sFontClose="</b></font>"
end if
'if ucase(left(strSQL,6))="SELECT" then
	rsTemp.open strSQL, sMDB
	if not rsTemp.eof then
		sHTML=sHTML&"<tr >"
		for each fld in rsTemp.fields
			sHTML=sHTML&"<td class="&sRowHeaderClass&">"&sFont&fld.name&sFontClose&"</td>"
			if bInline=1 then
				if isnumeric(rsTemp(0)) or fld.name="Hits per minute"  then
					sAlign="Align=right"
					sVal=formatnumber(fld.value,iDecimal_Places)
				else
					sVal=rsTemp(0)
					sAlign=""
				end if
				on error resume next
				sHTML=sHTML&"<td "&sAlign&">"&sLink&sFont&sVal&sFontClose&sLinkC&"</td>"
				on error goto 0
			end if		
		next
		sHTML=sHTML&"</tr><tr>"

		if bInline=0 then
			do until rsTemp.eof
				iCount=iCount+1
				for each fld in rsTemp.fields
					if request("RedirectPath")<>"" then
						sLink="<a href=""" & request("RedirectPath") &fld.name& "=" &rsTemp.fields(fld.name)& """>"
						sLinkC="</a>"
					end if
					if isnumeric(fld.value) or fld.name="Hits per minute" then
						sAlign="Align=right"
						sVal=formatnumber(fld.value,iDecimal_Places)
					else
						sVal=fld.value
						sAlign=""
					end if
					on error resume next
					sHTML=sHTML&"<td "&sAlign&">"&sLink&sFont&sVal&sFontClose&sLinkC&"</td>"
					on error goto 0
				next
				rsTemp.movenext
				sHTML=sHTML&"</tr><tr>"
			loop
		end if
		sHTML=sHTML&"</tr>"
		if bInline=0 then sHTML=sHTML&"</table>"
		DisplaySQL=sHTML
	else
		response.write "No records for that search!</br>"
	end if
	
'else
'	DisplaySQL="Must be a Select Statement for this function"
'end if
end function
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
Function RunProcedure(strSQL,bShowColumns)
dim rsRP
set rsRP=server.createobject("ADODB.recordset")
'x=rwe(strSQL)
RunProcedure=""
sFont=""       
iDecimal_Places=2
sFontClose=""
sRowHeaderClass=""
sLink=""
sHTML="<table class=""tblResults"">"
if bHighLight=1 then
	 sFont="<font color=red><b>"
	 sFontClose="</b></font>"
end if
'if ucase(left(strSQL,6))="SELECT" then
	'response.write strSQL
	'response.end
	i=0
	j=0
	k=1
	iPage=1
	iPgSize=2000
	if not request("pg")="" then iPage=request("pg")
	'x=rwe(strSQL)
	rsRP.open strSQL, sMDBA
	on error resume next
	beof=  rsRP.eof
	if err.number>0 then
		RunProcedure="<tr><td cospan=""20""><span class=""good"">"&"Empty Recordset Returned (Action complete)!</span></td></tr></table>"
		exit Function
	end if
	
	if not rsRP.eof then
		'x=rwbe(strSQL)
		'get record count
		'do until rsRP.eof
		'	i=i+1
		'	rsRP.movenext
		'loop
	else
		RunProcedure="<tr><td cospan=""20"">"&"Query has no results to return!</td></tr></table>"
		exit Function
	end if
	i=0
	rsRP.movefirst
	
	'function MakePages(sUrl,iCount,iPgFile,iPagesShow,iCurrPage,sHTM)
	sHTML=sHTML&"<tr><td cospan=""20"">"&MakePages(sScriptName&"?p="&request("pro"),i,iPgSize,10,request("pg"),"")&"</td></tr>"
	'do until rsRP.eof
	'	j=j+1
	'	if j>(iPage-1)*iPgSize then exit do
	'	rsRP.movenext
	'loop
	on error goto 0
	s=rsRP(0)
	if err.Number=0 then
		on error goto 0
		if bShowColumns then
			sHTML=sHTML&"<tr>"
			'x=rwb("here")
			for each fld in rsRP.fields
				sHTML=sHTML&"<td class=""qFieldNames"">"&sFont&fld.name&sFontClose&"</td>"
				if bInline=1 then
					if isnumeric(rsRP(0)) or fld.name="Hits per minute"  then
						sAlign="Align=right"
						sVal=formatnumber(fld.value,iDecimal_Places)
					else
						sVal=rsRP(0)
						sAlign=""
					end if
					'on error resume next
					if isnull(sVal) then sVal=""
					sHTML=sHTML&"<td "&sAlign&">"&sLink&sFont&sVal&sFontClose&sLinkC&"</td>"
					on error goto 0
				end if		
			next
			sHTML=sHTML&"</tr>"
		end if
		sHTML=sHTML&"<tr>"
		if bInline=0 then
			do until rsRP.eof
				iCount=iCount+1
				for each fld in rsRP.fields
					if request("RedirectPath")<>"" then
						sLink="<a href=""" & request("RedirectPath") &fld.name& "=" &rsRP.fields(fld.name)& """>"
						sLinkC="</a>"
					end if
					if isnumeric(fld.value) or fld.name="Hits per minute" then
						sAlign="Align=right"
						sVal=formatnumber(fld.value,iDecimal_Places)
					else
						sVal=fld.value
						sAlign=""
					end if
					if isnull(sVal) then sVal=""
					'on error resume next
					sHTML=sHTML&"<td "&sAlign&">"&sLink&sFont&sVal&sFontClose&sLinkC&"</td>"
					on error goto 0
				next
				if iCount>iPgSize then exit do
				rsRP.movenext
				sHTML=sHTML&"</tr><tr>"
				'x=rwb("here2")
			loop
		end if
		sHTML=sHTML&"</tr>"
	else
		'sHTML=sHTML& "<tr><td class=""good"">Query has been run sucessfully!</td></tr></table>"
	end if
	on error goto 0
'else
rsRP.close
set rsRP=nothing
'	DisplaySQL="Must be a Select Statement for this function"
'end if
sHTML=sHTML&"</table>"
RunProcedure=sHTML
end function
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
Function RunProcedureSingle(strSQL)
dim rsProSingle
set rsProSingle=server.createobject("ADODB.recordset")

RunProcedureSingle=""
sFont=""
sFontClose=""
sRowHeaderClass=""
sLink=""
sHTML="<table class=""tblResults"">"
if bHighLight=1 then
	 sFont="<font color=red><b>"
	 sFontClose="</b></font>"
end if
'x=rwe(strSQL)
rsProSingle.open strSQL, sMDB
'on error resume next
s=rsProSingle(0)
	for each fld in rsProSingle.fields
		sHTML=sHTML&"<tr >"
		sHTML=sHTML&"<td class=""qFieldNames"">"&sFont&fld.name&sFontClose&"</td>"
		if isnumeric(fld.value) then
			sAlign="Align=right"
			sVal=formatnumber(fld.value,iDecimal_Places)
		else
			sVal=fld.value
			sAlign=""
		end if
		sHTML=sHTML&"<td "&sAlign&">"&sLink&sFont&sVal&sFontClose&sLinkC&"</td>"
		sHTML=sHTML&"</tr><tr>"
	next
	
sHTML=sHTML&"</table>"
RunProcedureSingle=sHTML
rsProSingle.close
set rsProSingle=nothing
end function
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
function stripWhiteSpace(sText)
sText=replace(sText,chr(13),"")
sText=replace(sText,chr(10),"")
sText=replace(sText,chr(9),"")
sText=replace(sText," ","")

stripWhiteSpace=sText
end function

'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'

Function stripHTML(strHTML)
	if not isnull(strHTML) then
		strHTML=replace(strHTML,"</br>",chr(13))
		strHTML=replace(strHTML,"&nbsp;"," ")
		strHTML=replace(strHTML,". ",". "&chr(13))
		 
	'Strips the HTML tags from strHTML
	  Dim objRegExp, strOutput
	  Set objRegExp = New Regexp
	  objRegExp.IgnoreCase = True
	  objRegExp.Global = True
	  objRegExp.Pattern = "<(.|\n)+?>"
	  'Replace all HTML tag matches with the empty string
	  strOutput = objRegExp.Replace(strHTML, "")
	  'Replace all < and > with &lt; and &gt;
	  strOutput = Replace(strOutput, "<", "&lt;")
	  strOutput = Replace(strOutput, ">", "&gt;")
	  stripHTML = strOutput    'Return the value of strOutput
	  'response.write stripHTML
		'response.end
	  Set objRegExp = Nothing
	 else
	 	stripHTML=""
	 end if
End Function
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
function getPage(sURL)
sPage=sURL
xobj.Open "GET",sPage,false
xObj.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
xObj.send 
getPage=xobj.responseText
end function

function FnFindText(sSearchText,sFindText,sFindOffset,sEndText,sEndOffset,sRegExp)
'sPg is textString to search
'sFindText is Text to find
'sFindOffset is offset from Start text string can be - or +
'sEndText is the text where searchtext ends
'sEndOffset is the offset text where searchtext ends
'sRegExp=Numeric or Text, (Optional)
if sFindText="" or len(sFindText)=0 then 
	sText=sSearchText
else
	iSt=0
	iSt=instr(sSearchText,sFindText)+sFindOffset
	if iSt+sFindOffset<=0 then
		FnFindText="Error: Start minus offset error in FnFindText routine: functions.asp line 3588."
		exit function
	end if
	'x=ifa(iSt&":"&sSearchText&","&sEndText)
	iFn=instr(iSt+1,sSearchText,sEndText)+sEndOffset
	if sEndText="" or sEndText="vbcr" then
		'note this works for any vbcr encoded files but may not work well for large files where you must enter next text
		iFn=len(sSearchText)
	end if
	if iFn<iSt then 
		FnFindText=sSearchText'="Error: Finish Location is lower value than start FN:"&iFn&" ST:"&iSt&" error in FnFindText routine: Searching '"&sSearchText&"' for '"&sFindText&"' functions.asp line 3598."
		exit function
	end if
	if iFn=iSt then 
		if not iFn and iSt=0 then
			'in the case where iFn and iSt are the same asumption is the text is the same
			'change the finish to be be from the after where the first start varible was found.
			iFn=instr(iSt+1,sSearchText,sEndText)+sEndOffset
			sText=mid(sSearchText,iSt,iFn-iSt)
			FnFindText=sSearchText
			exit function
		end if
	end if
	sText=mid(sSearchText,iSt,iFn-iSt)
	if len(sRegExp)>0 then
		sText=RegRip(sText,sRegExp)
	end if
end if
FnFindText=sText
end function

function FnFindTextLIO(sSearchText,sFindText,sFindOffset,sEndText,sEndOffset,sRegExp)
'Same as fnFindText Except finds the last instance of text
'response.write sFindText
'response.end
iSt=1
do until instr(iSt,sSearchText,sFindText)=0
	iSt=instr(iSt,sSearchText,sFindText)+sFindOffset
	'response.write iSt&"</br>"
	iLoopCount=iLoopCount+1
	if iLoopCount>20 then 
		response.write "Loop count error  in function FnFindTextLIO"
		response.end
	end if
	iLastPos=iSt
loop
if iLastPos-sFindOffset<=0 then
	FnFindText=""
	exit function
end if
iFn=instr(iLastPos,sSearchText,sEndText)+sEndOffset
if iFn<=iSt then 
	FnFindText=""
	exit function
end if
sText=mid(sSearchText,iSt,iFn-iSt)
sText=RegRip(sText,sRegExp)
FnFindTextLIO=sText

end function

function RegRip(sVar,sRegExp)

if sRegExp<>"" then
	Dim objRegExpr 
	Set objRegExpr = New regexp
	objRegExpr.Pattern = sRegExp 
	objRegExpr.Global = True 
	objRegExpr.IgnoreCase = True	
	Set colMatches = objRegExpr.Execute(sVar)
	For Each objMatch in colMatches 
		sRegExpOutput=sRegExpOutput&objMatch.Value 
	Next
	Set colMatches = Nothing 
	Set objRegExpr = Nothing
	RegRip=sRegExpOutput
else
	RegRip=sVar
end if
end function


Function TrimPage(sPage,sFindText,bEndofText)
if instr(sPage,sFindText)>0 then
	if bEndofText=1 then
		iLenFindText=len(sFindText)-1
	else
		iLenFindText=-1
	end if
	sPage=right(sPage,len(sPage)-instr(sPage,sFindText)-iLenFindText)
else
	sPage=sPage
end if
TrimPage=sPage
end function

'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
Function WinHTTPPostRequest(URL, FormData, Boundary)
  Dim http 'As New MSXML2.XMLHTTP
	Const HTTPREQUEST_SETCREDENTIALS_FOR_SERVER = 0
	Const HTTPREQUEST_SETCREDENTIALS_FOR_PROXY = 1
  on error goto 0
  'Create XMLHTTP/ServerXMLHTTP/WinHttprequest object
  'You can use any of these three objects.
  Set http = CreateObject("WinHttp.WinHttprequest.5.1")
  http.Open "POST", URL, False
  http.setCredentials sTMUID, sTMPWD, HTTPREQUEST_SETCREDENTIALS_FOR_SERVER 
  http.setRequestHeader "Content-Type", "multipart/form-data; boundary=" + Boundary
  http.SetTimeouts 60000, 60000, 60000, 60000
  'Set Content-Type header
  'Send the form data To URL As POST binary request
  http.send FormData
  'response.write "OK:"
  'Get a result of the script which has received upload
  sPage = http.responseText
  'response.write sPage
  'response.end
  sPhotoID=mid(sPage,instr(sPage,"photo=")+6,9)
  'response.write sPhotoID & "</br>"
  'response.write instr(sPage,"photo=") & "</br>"
  'response.end
  if isNumeric(sPhotoID) then
  	WinHTTPPostRequest=sPhotoID
  else
  	WinHTTPPostRequest=sPage
	end if
End Function
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'Returns file contents As a binary data
Function GetFile(FileName)
  Dim Stream: Set Stream = CreateObject("ADODB.Stream")
  Stream.Type = 1 'Binary
  Stream.Open
  Stream.LoadFromFile FileName
  GetFile = Stream.Read
  Stream.Close
End Function
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'Converts OLE string To multibyte string
Function StringToMB(S)
  Dim I, B
  For I = 1 To Len(S)
    B = B & ChrB(Asc(Mid(S, I, 1)))
  Next
  StringToMB = B
End Function
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
Function  LogIn(sUID,sPWD) 
strPostData="email=" & Server.URLEncode(sUID)& "&password=" & sPWD & "&auto_login=true&attempts=0&submitted=Y&url=&login_attempts=0"
'response.write strPostData
xobj.Open "POST","https://secure.trademe.co.nz/Members/SecureLogin.aspx?session={EAE0FA78-690D-4BEF-8A0E-E7B05C0F9F48}",false
xObj.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
xObj.send strPostData
strRetval=xobj.responseText
if instr(xObj.responsetext,"Your Trade Me account is in debt.")>0 and bTestUpload=false then 
	response.write "You Need to Credit your TradeMe account before continuing.</br><tr><font color=red><b>Not enough Credit Error.....Uploader halted click <a href=http://www.trademe.co.nz/Payments/ChooseMethod.aspx>here</a> to add more credit." & "</b></font></br></tr>"
	response.end
end if
session("TM_Login")=session("dbo")
'x=rwe(strRetval)
end Function 
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
Function  LogInNoCheck(sUID,sPWD) 
strPostData="email=" & Server.URLEncode(sUID)& "&password=" & sPWD & "&auto_login=true&attempts=0&submitted=Y&url=&login_attempts=0"
'response.write strPostData
xobj.Open "POST","https://secure.trademe.co.nz/Members/SecureLogin.aspx?session={EAE0FA78-690D-4BEF-8A0E-E7B05C0F9F48}",false
xObj.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
xObj.send strPostData
strRetval=xobj.responseText
session("TM_Login")=session("dbo")
'response.write strRetval
end Function 
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
function LogOut(sTMUID,sTMPWD)
sTMURL="http://www.trademe.co.nz/Members/Login.aspx"
sPost="logout_auto=" & Server.URLEncode("Turn off auto-login")
xobj.Open "POST",sTMURL,false
xObj.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
xObj.send sPost
session("TM_Login")=0
end Function 


'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'

function ApplyOverlayLogo(sImagePath,sLogoOverlayPath,sNewFilePath,iLogoOverlayTop,iLogoOverlayLeft,iLogoOverlayTransparency,iLogoHeight,iLogoWidth)
'Used to apply overlay of one image on top of another
'on error Resume Next 

'response.write sImagePath&"</br>"&sLogoOverlayPath
set JpegImage = Server.CreateObject("Persits.Jpeg")
Set JpegOverLay = Server.CreateObject("Persits.Jpeg")
'x=rwe(sImagePath)
JpegImage.Open sImagePath
JpegOverLay.Open sLogoOverlayPath
JpegImage.PreserveAspectRatio = True
x=rwb(sImagePath)
x=rwb(sLogoOverlayPath)

'x=rwb(JpegImage.Height)
'x=rwb(JpegOverLay.Height)
if JpegImage.Height >600 then 
	JpegImage.Height=600 
end if
if JpegImage.width>800 then JpegImage.width=800 
if iLogoHeight>iLogoWidth then 
	if JpegImage.Height<>JpegOverLay.height then  JpegImage.Height=JpegOverLay.height
else 
	if JpegImage.width <>JpegOverLay.width then  JpegImage.width=JpegOverLay.width
end if
'x=rwb(JpegImage.Height)
'x=rwb(JpegOverLay.Height)
'Response.end
if err.Number<>0 then 
	ApplyOverlayLogo="path not found error"
	exit function 
end if
'on error goto 0

if iLogoOverlayTop="" then iLogoOverlayTop=0
if iLogoOverlayLeft="" or not isnumeric (iLogoOverlayLeft) then iLogoOverlayLeft =0
if iLogoOverlayTransparency="" or not isnumeric (iLogoOverlayTransparency) then iLogoOverlayTransparency =0
'response.end
iImageWidth=JpegImage.Width
iImageHeight=JpegImage.Height
bConstrainProportions=true 
if not isnumeric(iLogoWidth) then 
	if instr(iLogoWidth,"%") and IsNumeric(replace(iLogoWidth,"%","")) then 
		iLogoWidthRatio=cint(replace(iLogoWidth,"%",""))
		iLogoWidth = iImageWidth*iLogoWidthRatio/100 
	else
		if instr(iLogoWidth,"px") then iLogoWidth=Replace(iLogoWidth,"px") else iLogoWidth=0
	end if
	if instr(iLogoHeight,"%") and IsNumeric(replace(iLogoHeight,"%","")) then 
		iLogoHeightRatio=cint(replace(iLogoHeight,"%",""))
		iLogoHeight = iImageHeight *iLogoHeightRatio/100 
	else 
		if instr(iLogoHeight,"px") then iLogoHeight=Replace(iLogoHeight,"px") else iLogoHeight=0
	end if	
end if
if iLogoWidth>0 then 
	JpegOverlay.width=iLogoWidth
	if bConstrainProportions=true then 
		iLogoHeight=iLogoHeight * iLogoWidth / JpegOverlay.OriginalWidth
	end if
end if
if iLogoHeight>0 then 
	JpegOverlay.Height =iLogoHeight
	if bConstrainProportions=true then 
		iLogoWidth = iLogoWidth * iLogoHeight / JpegOverlay.OriginalHeight
	end if
end if

SavePath = sNewFilePath
iXPos=(JpegImage.OriginalWidth-JpegOverlay.OriginalWidth)/2
if iXPos<0 then iXPos=0
iYPos=(JpegImage.OriginalHeight-JpegOverlay.OriginalHeight)/2
if iXPos<0 then iXPos=0
'on error resume next
'response.write "JpegImage.Canvas.DrawImage "&iLogoOverlayTop&","&iLogoOverlayLeft&",JpegOverlay,"&iLogoOverlayTransparency/100

JpegImage.ToRGB

JpegImage.Canvas.DrawPNG iLogoOverlayTop,iLogoOverlayLeft, sLogoOverlayPath
'on error goto 0
'JpegImage.Canvas.DrawImage 10, 10, JpegOverlay
ApplyOverlayLogo= "</br>DrawImage applied to "&sImagePath &"</br>"
'x=rwe(SavePath)
JpegImage.Save SavePath
'response.write "<tr><td colspan=""10"">New Logo File is <a href="""&replace(replace(SavePath,sSitePath,sSiteURL),"\","/")&""">here</a></br>"
JpegImage.close
JpegOverlay.Close
set JpegImage =nothing 
set JpegOverlay =nothing 
ApplyOverlayLogo=ApplyOverlayLogo& SavePath & " Saved!" & "</br>"
on error goto 0
'response.end 
'x=rwe(SavePath)
end function 
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
function CheckExistsFile(sFilePath)
dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
if fso.FileExists(sFilePath) then
	CheckExistsFile=true
else
	CheckExistsFile=false
end if
set fso=nothing
end function

function ReNameFile(sFilePathtoRename,sNewFileNameLocation)
dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
if fso.FileExists(sFilePathtoRename) then
	if not fso.FileExists(sNewFileNameLocation) then
		fso.MoveFile sFilePathtoRename, sNewFileNameLocation
		ReNameFile=sFilePathtoRename &" moved to "& sNewFileNameLocation
	else
		fso.DeleteFile sFilePathtoRename,1
		ReNameFile=sNewFileNameLocation &" already exists "& sFilePathtoRename & "Deleted"
	end if
else
	ReNameFile="No Update for "&sFilePathtoRename&": file not found"
end if
set fso=nothing
end function

function UploadTMImages(sPn)
dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
for i=1 to 10
	if i=1 then sPhoto="" else sPhoto="-Photo"&i
	path=sSitePath&"\Images\DB"&idb&"\Stock\Original\"&sPn&sPhoto&".jpg"
	if not fs.FileExists(path) then path=sSitePath&"\Images\DB"&idb&"\Stock\X-Large\"&sPn&sPhoto&".jpg"
	if not fs.FileExists(path) then path=sSitePath&"\Images\DB"&idb&"\Stock\Large\"&sPn&sPhoto&".jpg"
	'x=rwe(path)
	sOutputPath=sSitePath&"\Images\DB"&idb&"\Stock\tm_photos\" & sPn & sPhoto&".jpg"
	sLink="<a href="""&sSiteURL&"/images/db"&idb&"/Stock/tm_photos/"&sPn&".jpg"&""">link "
	'x=rwb("File:"&path&":"&fs.FileExists(path))
	'x=rwb("File:"&sOutputPath&":"&fs.FileExists(sOutputPath))
	if fs.FileExists(path) then	
		if not fs.FileExists(sOutputPath) then
			x= ApplyOverlayLogo(path,sLogoOverlayPath,sOutputPath ,iLogo_Top,iLogo_Left,iLogo_Transparancy,sLogo_Height,sLogo_Width)
		end if
		xobj.Open "GET","https://www.laptopbattery.co.nz/admin/trademe/addPhoto.aspx?p="&sOutputPath&"&u="&sTMID&"&d="&idb,false
		xObj.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
		xObj.send 
		strRetval=xobj.responseText
		x=rwb(strRetval)
		if instr(strRetval,"Success")>0 then
			UploadTMImage=fnFindText(strRetval,"Id>",3,"</PhotoId>",0,"")
		else
			UploadTMImage=0
		end if
		sSQL="P_up_tm_photos '"&sPn&"',"&i&","&UploadTMImage&",'"&sOutputPath&"',"&idb
		'x=rwb(sSQL)
		x=openRSA(sSQL)
		x=closeRSA()
	end if
next
set fs=nothing
end function 


'******************* upload - end
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'

Function CheckFreeListings ()
sTMURL="http://www.trademe.co.nz/MyTradeMe/WeeklySummary.aspx"
sPost=""
xobj.Open "GET",sTMURL,false
xObj.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
xObj.send 
sPage=xobj.responseText
iListingsAllowance=FnFindText(sPage,"you can list ",len("you can list ")," concurrent ",0,"")
CheckFreeListings=iListingsAllowance
end Function
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
function CheckBalance()
'sTMURL="http://www.trademe.co.nz/MyTradeMe/Default.aspx"
'sPost=""
'xobj.Open "GET",sTMURL,false
'xObj.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
'xObj.send 
'sPage=xobj.responseText
'sBal=FnFindText(sPage,"Balance: <span style=""color: #339900;""><b>$",len("Balance: <span style=""color: #339900;""><b>$"),"</b>",0,"")
CheckBalance=-14.07
end Function
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
Function formatDayHour(sDate)
if not isdate(sDate) then 
	formatDayHour=sDate
else 
	formatDayHour=WeekdayName (weekday (sDate))&" "&Hour(sDate)&":"&right("0"&Minute(sDate),2)&"."&right("0"&Second(sDate),2)
end if
end function 

'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
'**********************************************************************************************'
Function Upload_Photo(iPhoto)
sTMURL="http://www.trademe.co.nz/Sell/UploadPhotoComplete.aspx?photo=" &iPhoto
'response.write sTMURL&"</br>"
'on error resume next
sPost=""
xobj.Open "POST",sTMURL,false
xObj.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
xObj.send sPost
strRetval=xObj.ResponseText
'on error goto 0
'response.write strRetval
if instr(strRetval,"You have uploaded the maximum number of photos.")>0 then 
		response.write "</br><font color=red>Maximum Photos Exceeded Error Please check before continuing...</font></br>"
	response.end
end if

if instr(strRetval,"You cannot add another photo to this")>0 then 
	response.write sTMURL
	response.write "</br><font color=red>You cannot add another photo to this error...please debug</font></br>"
	response.end
end if

if session("debug")=true then
	response.write "<tr class=""tmstart""><td style=""width:800px"" coslpan=""10"">------Showing Post for UPLOAD_PHOTO.ASPX Page START </td></tr>"
	response.write "<tr><td class=""tmpost"">"
	response.write strRetval
	response.write "</td></tr>"
	response.write "<tr class=""tmfinish""><td style=""width:800px"" coslpan=""10"">------Showing Post for UPLOAD_PHOTO.ASPX Page FINISH </td></tr>"
end if
end Function 

function StockModify(iQty)
if iQty=>10 then sq="10+"
if iQty<10 and iQty>4 then sq="5+"
if iQty<5 then sq="1+"
if iQty=1 then sq="1 only!"
if iQty=0 then sq="SOLD OUT!"
StockModify=sq
end function

function StockQuantityColors(iQty)
if iQty>10 then sq="lots"
if iQty<10 and iQty>4 then sq="gettingLow"
if iQty<5 then sq="almostGone"
if iQty=0 then sq="soldOut"
StockQuantityColors=sq
end function


function DrawQuery(qID,iTop,iLeft,iWidth,iHeight)
'Point in case is need to draw a query in HTML
'set varibles for where the query position and szie etc
'Properties, Fields, Feild_name, field_width,Field_color ect
if iHeight=0 then iHeight="auto"
if iWidth=0 then iWidth="auto"
sDebug=sDebug& "running QID:"&qID&"</br>"
dim sDiv
sDiv=""
dim rs
set rs=server.createobject("adodb.recordset")

rs.open "P_Sel_QueryForDisplay " &qID,sMDBA
'getQuery Name and SQL
if rs.eof then
	DrawQuery="Invalid Query Specified! :"&qID&"</br>"
	exit function
end if
sQuerySQL=rs("sql")
sQueryName=rs("name")
bQueryshowColumns=rs("show_columns")
sFooter=rs("footer")
sRowClass=rs("class")

'response.end

sQuerySQL=ReplaceVar(sQuerySQL,0,0,1,"")
'x=ifa(sQuerySQL)
if request("p1")<>"" then
	'if instr(sQuerySQL," ")>0 then sQuerySQL=sQuerySQL&","
	if instr(sQuerySQL,"'P1'")>0 then 
		sQuerySQL=replace(sQuerySQL,"'P1'","'"&request("p1")&"'")
	else
		sQuerySQL=sQuerySQL&" '"&request("p1")&"'"
	end if
end if
if request("p2")<>"" then
	if instr(sQuerySQL,"'P2'")>0 then sQuerySQL=replace(sQuerySQL,"'P2'","'"&request("p2")&"'")
	sQuerySQL=sQuerySQL&",'"&request("p2")&"'"
end if
if request("p3")<>"" then
	if instr(sQuerySQL,"'P3'")>0 then sQuerySQL=replace(sQuerySQL,"'P3'","'"&request("p3")&"'")
	'sQuerySQL=sQuerySQL&",'"&request("p3")&"'"
end if
if instr(sQuerySQL,"'P1'")>0 then sQuerySQL=replace(sQuerySQL,"'P1'","''")
if instr(sQuerySQL,"'P2'")>0 then sQuerySQL=replace(sQuerySQL,"'P2'","''")
if instr(sQuerySQL,"'P3'")>0 then sQuerySQL=replace(sQuerySQL,"'P3'","''")
rs.close
on error resume next 
rs.open sQuerySQL,sMDBA

if session("debug")=true then x=rwSQL(sQuerySQL)
sDebug=sDebug&"running SQL:<span class=""good"">"&sQuerySQL&"</span></br>"
if err.number<>0 then
response.write "There is an error with procedure number <a href=""/admin/item.asp?t=85&id="&qID&""">"&qID&".<span class=""attention"">"&err.description&"</span></br></br>"
if session("staff_id")=1 then
	x=rwe(sQuerySQL)
end If
response.end
end if
on error goto 0 
sDiv=sDiv&"<div id=""QID"&qID&""" style=""float:left;top:"&iTop&"px;left:"&iLeft&"px;width:"&iWidth&"px;height:"&iHeight&"px;overflow:none;padding:0px;margin:10px;"">"&vbcrlf
sDiv=sDiv&"<table class=""qBody"" height=""95%"">"&vbcrlf
sDiv=sDiv&"<tr class=""qHead"">"&vbcrlf
sDiv=sDiv&"<td colspan="""&rs.Fields.count&"""><div style=""text-align:left;float:left;"">"&sQueryName&"</div>"
if session("staff_id")<>"" then
	sDiv=sDiv&"<div style=""text-align:right;float: right;""><a href=""/admin/dashboardQ.asp?q="&qID&""">expand +</a></div>"
end if
sDiv=sDiv&"</td>"&vbcrlf
sDiv=sDiv&"</tr>"&vbcrlf
sDiv=sDiv&"<tr><td colspan="""&rs.Fields.count&" style=""padding:0px;"">"

if not iWidth="auto" then
	sDVwidth=iWidth-6
else
	sDVwidth=iWidth
end if
if not iHeight="auto" then
	sDVHeight=iHeight-30
else
	sDVHeight=iHeight
end if
sDiv=sDiv&"<div style=""width:"&sDVwidth&"px;height:"&sDVHeight&"px;overflow:auto;"">"
sDiv=sDiv&"<table style=""width:100%"">"
iRow=0
iFld=0
if bQueryshowColumns then
	sDiv=sDiv&"<tr class=""qFieldNames"">"&vbcrlf
	for each fld in rs.fields
		'field sytle should be calucalted from field_lookup width property in pixels
		if iFld mod 2 = 0 then
			sDiv=sDiv&"<td>"&fld.name&"</td>"&vbcrlf
		end if
		iFld=iFld+1
	next
	sDiv=sDiv&"</tr>"&vbcrlf	
end if
do until rs.eof
	iRow=iRow+1
	if iRow mod 2=1 then
		sRowClass="light_blue_row"
	ELSE
		sRowClass=""
	end if
	sDiv=sDiv&"<tr class="""&sRowClass&""">"&vbcrlf
	for iFld=0 to rs.fields.count-1 step 2
	if iFld=0 then sHasRowClass="class="""&sRowClass&"""" else sHasRowClass=""
		'field sytle should be calucalted from field_lookup width property in pixels
		if left(rs(iFld+1),6)="class=" then
			sStyling=rs(iFld+1)
		else
			sStyling="style="""&rs(iFld+1)&""""
		end if
		sDiv=sDiv&"<td "&sStyling&" "&sHasRowClass&">"&rs(iFld)&"</td>"&vbcrlf
	next
	sDiv=sDiv&"</tr>"&vbcrlf
	rs.movenext
loop
if sFooter<>"" then
	sDiv=sDiv&"<tr><td colspan="""&rs.Fields.count&""" style="""">"&sFooter&"</td></tr>"
end if
sDiv=sDiv&"</table></div>"
sDiv=sDiv&"</td></tr></table>"&vbcrlf
sDiv=sDiv&"</div>"&vbcrlf
DrawQuery=sDiv
DrawQuery=DrawQuery&"<script type=""text/javascript"" src=""/scripts/drawQuery.js""></script>"&vbcrlf

rs.close
set rs=nothing
end function


function DrawQueryAdvanced(qID,iTop,iLeft,sWidth,sHeight,sParams,bShowHeader)
'Point in case is need to draw a query in HTML
'set varibles for where the query position and szie etc
'Properties, Fields, Feild_name, field_width,Field_color ect
bShowSeperator=1
sDebug=sDebug& "running QID:"&qID&"</br>"
dim sDiv
sDiv=""
dim rs
set rs=server.createobject("adodb.recordset")
'response.write "P_Sel_QueryForDisplay " &qID
rs.open "P_Sel_QueryForDisplay " &qID,sMDB
'getQuery Name and SQL
if rs.eof then
	DrawQueryAdvanced="Invalid Query Specified! :"&qID&"</br>"
	exit function
end if
sQuerySQL=rs("sql")
sQueryName=rs("name")
bQueryshowColumns=rs("show_columns")
sFooter=rs("footer")
sRowClass=rs("class")
'response.end
sQuerySQL=ReplaceVar(sQuerySQL,0,0,1,"")
if sParams<>"" then sQuerySQL=sQuerySQL&" "&sParams
rs.close
'on error resume next 
'response.write sQuerySQL
'response.end
rs.open sQuerySQL,sMDBA
'rs.open "P_Email_MessagesByFolder 1,0,'',1",sMDB
sDebug=sDebug&"running SQL:<span class=""good"">"&sQuerySQL&"</span></br>"
if session("debug")=true then x=rwSQL(sDebug)
if err.number<>0 then
	sDebug=sDebug&"running SQL:<span class=""good"">"&sQuerySQL&"</span></br>"
	sDebug=sDebug&"There is an error with procedure number <a href=""/admin/item.asp?t=85&id="&qID&""">"&qID&".<span class=""attention""></br></br>"&err.description&"</span>"
	response.write sDebug
	response.end 
end if
on error goto 0 
if iTop=0 and iLeft=0 then
	sPosType="relative"
Else
	sPosType="absolute"
end if
sDiv=sDiv&"<div id=""QID"&qID&""" style=""position:"&sPosType&";top:"&iTop&"px;left:"&iLeft&"px;width:"&sWidth&";height:"&sHeight&";overflow:none;padding:0px;"">"&vbcrlf
sDiv=sDiv&"<table class=""etbl"&qID&""" style=""width:100%;height:100%;text-align:top;"">"&vbcrlf
iRow=0
iFld=0
if not rs.eof then
	
do until rs.eof
	
		iRow=iRow+1
		if iRow=1 then
			'write header column and query titleactivated
			if bShowHead=1 then
				sDiv=sDiv&"<tr class=""qHead"">"&vbcrlf
				sDiv=sDiv&"<td colspan="""&rs.Fields.count&""">"&sQueryName&"</td>"&vbcrlf
				sDiv=sDiv&"</tr>"&vbcrlf
			end if
			sDiv=sDiv&"<tr><td valign=""top"" colspan="""&rs.Fields.count&""" style=""padding:0px;""><div style=""width:"&iWidth-6&"px;height:"&iHeight-30&"px;overflow:auto;""><table style=""width:100%"">"
			if bQueryshowColumns then
				sDiv=sDiv&"<tr class=""qFieldNames"">"&vbcrlf
				for each fld in rs.fields
					'field sytle should be calucalted from field_lookup width property in pixels
					if iFld mod 2 = 0 then
						iFldCount=iFldCount+1
						sfName=fld.name
						if instr(sfName,"_pram")>0 then
							'add PamaterCode in here
							stype=right(sfName,1)
							iParam=left(right(sfName,2),1)
							sfName=left(sfName,instr(sfName,"_pram")-1)
							if sType="b" then
								'drop down list with All, Yes, No
								sDiv=sDiv&"<td>"&sfName&"</br><select id="""&sfName&""" onchange=""cps('p"&iParam&"',this.value)"">"
								sDiv=sDiv&"<option value=""-1"""
								if request("p"&iParam)="-1" then sDiv=sDiv&" selected"
								sDiv=sDiv&">All</option>"
								sDiv=sDiv&"<option value=""0"""
								if request("p"&iParam)="0" then sDiv=sDiv&" selected"
								sDiv=sDiv&">false</option>"
								sDiv=sDiv&"<option value=""1"""
								if request("p"&iParam)="1" then sDiv=sDiv&" selected"
								sDiv=sDiv&">true</option>"
								sDiv=sDiv&"</select></td>"
							end if
						Else
							if fld.name<>"rownumber" then sDiv=sDiv&"<td>"&fld.name&"</td>"&vbcrlf
						end if
					end if
					iFld=iFld+1
				next
				sDiv=sDiv&"</tr>"&vbcrlf	
			end if
		else
			sDiv=sDiv&"<tr id=""""tr"&iRow-1&sHTM&""">"&vbcrlf
			for iFld=0 to rs.fields.count-1 step 2
			if iFld=0 then sHasRowClass="class="""&sRowClass&"""" else sHasRowClass=""
				if sHasRowClass="" and bShowSeperator=1 then
					if iRow mod 2 = 0 then
						sHasRowClass="class=""light_blue_row"""
					else
						sHasRowClass=""
					end if
				end if 
				'field sytle should be calucalted from field_lookup width property in pixels
				if left(rs(iFld+1),6)="class=" then
					sStyling=rs(iFld+1)
				else
					sStyling="style="""&rs(iFld+1)&""""
				end if
				 if rs.fields(iFld).name<>"rownumber" then sDiv=sDiv&"<td valign=""top"" "&sStyling&" "&sHasRowClass&">"&rs(iFld)&"</td>"&vbcrlf
			next
			sDiv=sDiv&"</tr>"&vbcrlf
			rs.movenext
		end if
	loop
else
	sDiv=sDiv&"<tr><td colspan="""&rs.Fields.count&""" style="""">No results for this query...</td></tr>"
end if
if sFooter<>"" then
	sDiv=sDiv&"<tr><td colspan="""&rs.Fields.count&""" style="""">"&sFooter&"</td></tr>"
end if
sDiv=sDiv&"</table></div></td></tr></table>"&vbcrlf
sDiv=sDiv&"</div>"&vbcrlf
DrawQueryAdvanced=sDiv
rs.close
set rs=nothing
DrawQueryAdvanced=sDiv
DrawQueryAdvanced=DrawQueryAdvanced&"<script type=""text/javascript"" src=""/scripts/drawQuery.js""></script>"&vbcrlf

end function

Function ReplaceVar(s,iTbl,iKeyID,sDefault,sType)
	sVOpen="[["
	sVClose="]]"
	sSQLTemp=ucase(s)
	bSQLFail=false 
	'x=rwe(s)
	set rs=server.CreateObject("ADODB.recordset")
	'sPg=sPg&(instr(s,sVOpen))&"</br>"
	'x=rwe(S)
	if instr(s,sVOpen)>0 THEN 
		dim aVar(10)
		i=0
		'x=rwe(S)
		do until instr(sSQLTemp,sVOpen)=0
			i=i+1
			'x=rwb(FnFindText(sSQLTemp,sVOpen,2,sVClose,0,""))
			aVar(i)=FnFindText(sSQLTemp,sVOpen,2,sVClose,0,"")
			sSQLTemp=StartFromText(sSQLTemp,sVClose,1)
			'x=rwe(sSQLTemp)
		loop
		'x=rwe(aVar(1))
		'response.end
		sNewSQL=ucase(s)
		for j=1 to i
			'replace variables with the appropiate value
			if left(aVar(j),2)="S_" Then 
				'must be a session variable
				sVar=right(aVar(j),len(aVar(j))-2)
				SELECT Case sVar
					Case "DB_OWNER"
					 sNewSQL=replace(sNewSQL,sVOpen&aVar(j)&sVClose,iDB)			 
					Case "STAFF_ID"
					 sNewSQL=replace(sNewSQL,sVOpen&aVar(j)&sVClose,session("STAFF_ID"))
					Case "IP_ADDRESS"
					 sNewSQL=replace(sNewSQL,sVOpen&aVar(j)&sVClose,session("IP_Address"))
					Case "ID_ACCT"
					 	sNewSQL=replace(sNewSQL,sVOpen&aVar(j)&sVClose,session("eAccount"))
					Case "EACCOUNT"
					 sNewSQL=replace(sNewSQL,sVOpen&aVar(j)&sVClose,session("eAccount"))
					Case "STAFF_SEC_LEVEL"
					 	sNewSQL=replace(sNewSQL,sVOpen&aVar(j)&sVClose,session("STAFF_SEC_LEVEL"))
					Case "SEC_LEVEL"
					 	sNewSQL=replace(sNewSQL,sVOpen&aVar(j)&sVClose,session("SEC_LEVEL"))
					Case "CU_ID"
					 	sNewSQL=replace(sNewSQL,sVOpen&aVar(j)&sVClose,session("cu_id"))
					Case "OR_CU_ID"
					 	sNewSQL=replace(sNewSQL,sVOpen&aVar(j)&sVClose,session("or_cu_id"))
					Case "CUS_SYS"
					 	sNewSQL=replace(sNewSQL,sVOpen&aVar(j)&sVClose,session("cus_sys"))
					Case "MY_ORDER_NO_DELIVERY"
					 	sNewSQL=replace(sNewSQL,sVOpen&aVar(j)&sVClose,session("MY_ORDER_NO_DELIVERY"))
					Case else 
					response.Write("unknown variable used in SQL:"&sVar)
					response.end					'sNewSQL=replace(s,sVOpen&aVar(j)&sVClose,session(sVar))
				end select
			end if
			if left(aVar(j),2)="R_" Then 
				'x=RWE(s)
				'response.end
				'must be a recordset variable
				
				sSQ="P_Sel_VarforReplace "&iTbl &","&iKeyID&",'"&right (aVar(j),Len(aVar(j))-2)&"',"&iDB 
				'x=rwe(sSQ)
				on error resume Next
				rs.open sSQ,SMDB
				if not err.number=0 then
					x=rwe( "Error Occured with SQL: "&sSQ&"</br>Please fix before proceding.")
				end if
				on error goto 0
				
				'x=rwe(rs(0))
				sExec=rs(0)
				rs.close 
				rs.open sExec,sMDB
				if rs.eof  THEN 
					'response.write "nothing to replace error."
					'x=rw(sSQ)
					'response.End 
					bSQLFail=true
					select case sType
					case "datetime", "nvarchar","nchar","varchar"
						sDefaultForType=""
					case else
						sDefaultForType=0
					end Select
					if sDefault="" then 
						sFailSQL="Select "&sDefaultForType&",''"
					else 
						sFailSQL="Select "&sDefault&","&sDefault&" "
					end if
				else 
					sValReplace=rs(0)
				end if
				if isnull(sValReplace) or sValReplace="" THEN sValReplace=""
				rs.close 
				sNewSQL=replace(sNewSQL,sVOpen&aVar(j)&sVClose,sValReplace)
				'response.write sVOpen&aVar(j)&sVClose&" (replaced:with"&sValReplace&")</br> "
				'response.write  sNewSQL&"</br>"
			end if		
		next
		ReplaceVar =sNewSQL 
		if bSQLFail=true Then ReplaceVar=sFailSQL
		'response.write ReplaceVar&" (replaced:)</br> "
		
	else 
		'no varibles found pass back original string
		ReplaceVar=s
	end if
	'ok take the string and check for varible definitions
				'response.Write(sNewSQL)
				'response.end
end function

function timestamp(dTime,istart)
if istart=1 then timestamp="query run at :"&FormatDateTime( Now(),3 )&"</br>"
if istart=2 then timestamp="query finished at :"&FormatDateTime( Now(),3 )&"</br>"
end function

function StartFromText(sFullString,sFindString,bEndOfText)
iMoveTo=instr(sFullString,sFindString)
if iMoveTo=0 then
	StartFromText=sFullString
else
	if bEndOfText=1 then iSub=len(sFindString)-1 else iSub=-1
	StartFromText=right(sFullString,len(sFullString)-iMoveTo-iSub)
end if
end function

function TMRemovePhoto(iPhotoID)
if iPhotoID="" then
	exit function
end if
sPG=TrimPage(sPG,"/Sell/RemovePhoto.aspx?photoid="&iPhotoID,1)
sLink="http://www.trademe.co.nz/Sell/RemovePhoto.aspx?photoid="&iPhotoID
x=RW(sLink)
xobj.Open "GET",sLink,false
xObj.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
xObj.send
if session("debug")=true then 
	response.write "removing previous photos: "&sLink &"</br>"
end if
end function

function TMremovePhotos()
	'get this page
	dim iloop
	iloop=0
	'http://www.trademe.co.nz/Sell/Photos.aspx?return_path=/Sell/Confirm.aspx
	sLink="http://www.trademe.co.nz/Sell/Photos.aspx?return_path=/Sell/Confirm.aspx"
	xobj.Open "GET",sLink,false
	xObj.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
	xObj.send
	sPG=xobj.responseText
	'response.Write sPG
	'response.end
	do until instr(sPG,"/Sell/RemovePhoto.aspx?photoid=")=0
		iPhotoID=FnFindText(sPG,"/Sell/RemovePhoto.aspx?photoid=",len("/Sell/RemovePhoto.aspx?photoid="),"""",0,"")  
		TMRemovePhoto(iPhotoID)
		'response.write xObj.ResponseText
		iloop=iLoop+1
		if iLoop>20 then exit do
	loop
	TMremovePhotos=iloop
end function

function ConvertToHTML_Input(sName,sValue,sFieldType,sRequired,iSize,sScriptinline,sReadOnly,sInputClass,bShown,bFile,sKey,iKeyID,ifl_id,bSearch,bAutoSave)
bHelp=true
iTabIndex=iTabIndex+1
sJavaS=""
if bHelp then 
	if bAutoSave then 
			sJavaS=" onblur=""hideHelp('help"&sName&"');ufkd("&iKeyID&","&ifl_id&",event,'b')"""
	ELSE
		sJavaS=" onblur=""hideHelp('help"&sName&"');"""	
	end if
ELSE
	if bAutoSave then sJavaS=" onblur=""ufkd("&iKeyID&","&ifl_id&",event,'b')"""
end if
if ifl_id=360 then
    sJavaS=" onblur=""hideHelp('help"&sName&"');$('#srCart').hide(250);"""
end if

if bSearch=true then 
    if ifl_id=360 then 
        sJavaS=sJavaS&" onKeyUp=""PhotoSearch(event,this,'srCart')"" "
    else
        sJavaS=sJavaS&" onKeyUp=""UpdateCombo(this.id,this.value,this.id,0,event)"" "
    end if
end if 
if bHelp=true then  sJavaS=sJavaS&" onfocus=""showHelp('help"&sName&"')"" "
sJavaS=sJavaS&" "&sScriptinline

select case sFieldType
case "bit"
	sType="Checkbox"
	intFieldLength=0
	intMaxFieldLength=0
	HTML_Value=" Value=""true"""
	if sValue=true or sValue="true" then 
		sChecked=" checked"
	else
		sChecked=""
	end if
	if sReadOnly=" readonly" then sReadOnly=" disabled"
	sonclick="onclick=""this.focus()"""
case "datetime","date","smalldatetime","datetime2","datetimeoffset"
	'all possible formats from SQLServer
	'time						12:35:29. 1234567
	'date						2007-05-08
	'smalldatetime	2007-05-08 12:35:00
	'datetime 			2007-05-08 12:35:29.123
	'datetime2 			2007-05-08 12:35:29. 1234567
	'datetimeoffset 2007-05-08 12:35:29.1234567 +12:15
	sType="Text"
	intMaxFieldLength=11
	intFieldLength=15
	if sFieldType="date" then
		intMaxFieldLength=11
		intFieldLength=15
		sDate=CvbShortdateTime(sValue,false)
	end if
	if sFieldType="datetime" then
		intMaxFieldLength=24
		intFieldLength=28
		sDate=CvbShortdateTime(sValue,true)
	end if	
	if sFieldType="smalldatetime" then
		intMaxFieldLength=19
		intFieldLength=23
		sDate=CvbShortdateTime(sValue,true)
	end if			
	if sFieldType="datetime2" then
		intMaxFieldLength=28
		intFieldLength=32
		sDate=CvbShortdateTime(sValue,true)
	end if	
	if sFieldType="datetime2" then
		intMaxFieldLength=35
		intFieldLength=39
		sDate=CvbShortdateTime(sValue,true)
	end if	
	HTML_Value=" Value='" & sDate & "'"
	'response.write sValue
	'response.end
	sInputClass="fldDate"
	sonclick="onclick=""showCalendar(this.id)"""
case "nvarchar","nchar","varchar"
	sType="Text"
	if instr(ucase(sName),"PASSWORD")>0 and session("staff_sec_Level")<4 then sType="Password"
	intMaxFieldLength=iSize
	intFieldLength=iSize+20
	if intFieldLength>100 then intFieldLength=80
	if iSize>100  and bShown=true then
		iCols=65
		iRows=2
		if 	iSize>200 then
				iCols=65
				iRows=3
		end if
		if 	iSize>500 then
				iCols=65
				iRows=5
		end if
		if 	iSize>1000 then
				iCols=65
				iRows=6
		end if
		if 	iSize>2000 then
				iCols=65
				iRows=6
		end if			
		'iRows=cint(iSize/80)
		if iRows>15 then iRows=15
		ConvertToHTML_Input="<textarea tabindex="""&iTabIndex&""" cols="""&iCols&""" rows="&iRows&" name=""" & sName & """ id=""" & sName & """ " & sReadonly & sJavaS & ">" & sValue & "</textarea>" 
	end if
	if isnull(sValue) then sValue=""
	'sValue=replace(sValue,"""","''")
	if instr(sValue,"""")>0 and instr(sValue,"'")=0 then
		HTML_Value=" Value='" & sValue & "'"
		
	else
		if instr(sValue,"""")>0 then sValue=replace(sValue,"""","")
		HTML_Value=" Value=""" & sValue & """"
	end if
case else
	sType="Text"
	intMaxFieldLength=iSize+3
	intFieldLength=cdbl(iSize*1)
	if isnull(sValue) then sValue=""
	if instr(sValue,"""")>0 and instr(sValue,"'")=0 then
		HTML_Value=" Value='" & sValue & "'"
	else
		if instr(sValue,"""")>0 then sValue=replace(sValue,"""","")
		HTML_Value=" Value=""" & sValue & """"
	end if
end Select
if bShown=false then
	sType="Hidden"
end if
if intMaxFieldLength>0 then smaxlength=" maxlength=""" & intMaxFieldLength & """ "
if intFieldLength>0 then sfldlength=" size=""" & intFieldLength & """ "
if sInputClass<>"" then sClass=" class=" & sInputClass 
if ConvertToHTML_Input="" then ConvertToHTML_Input="<input type=""" & sType & """ "& sClass & " tabindex="""&iTabIndex&""" size=""" & intFieldLength & """ " & HTML_Value & smaxlength & sfldlength & " name=""" & sName & """ id=""" & sName & """ " & sReadonly & sChecked & sJavaS&">"

if bFile=true then 
	'if iKeyID=0 then
		'quick check to see fi any files exist at the location in question
		sRootDir=sSitePath&"\DB_Owner_files\DB"&iDB&"\"
		sD = sTable &"\"
		sD=sD&iKeyID &"\"
		sD=sD&sName
		URLLink=sSiteURL&"/DB_Owner_Files/DB"&iDB&"/"&replace(sD,"\","/")&"/"
		'response.write sRootDir&sD
		'response.end
		call OpenDB()
		'rsTemp.open 
		Call CloseDB()
		sCI=ShowFileList(sRootDir&sD,URLLink)
		'response.write sRootDir&sD
		'response.end	
		ConvertToHTML_Input="<a href=""db_owner_files/DB"&iDB&"/"&sTable&"/"&sKey&"/"&sValue
		sCI=sCI&"<img src=""/images/fileIcons/attach.gif"" border=""0""><a href=""#"" onclick=""window.open('"&sSecureURL&"/aspcomponents/Upload/AddFiles.asp?folder="&sTable&"&id="&iKeyID&"&field="&sName&"','addFiles','menubar=1,resizable=1,width=350,height=250,left=50,top=50')""> Attach a file</a>"
		ConvertToHTML_Input=sCI
	'end if
end if
end function

Function ShowFileList(folderspec,URLLink)
	'x=rwe("")
  Dim fso, f, f1, fc, s
  sErrMsg=""
  Set fso = CreateObject("Scripting.FileSystemObject")
  if not fso.folderExists(folderspec) then 
  	'response.write folderspec & " Not found!"
  	Set fso = Nothing
  	exit Function
  end if
  Set f = fso.GetFolder(folderspec)
  Set fc = f.Files
	set rsFiles=server.CreateObject("ADODB.recordset")
  iFiles=0
  For Each f1 in fc
  	'response.write "P_Sel_FileIDByLocation '"&folderspec&"\"&f1.name&"'</br>"
  	rsFiles.open "P_Sel_FileIDByLocation '"&folderspec&"\"&f1.name&"'",sMDB
  	if rsFiles.eof then 
  		sErrMsg=sErrMsg& "<tr><td colspan=""2"">no matching file '"&folderspec&"\"&f1.name&"' found in Database.  Please get the administrator to check</td></tr>"
 		else
	  	iFileID=rsFiles("F_ID")
	  	iFileSize=formatNumber(rsFiles("F_Disk_Size"),0)
		  if iFileSize<1000000 then sDiskSize=iFileSize/1000 & " KB" else sDiskSize=iFileSize/1000000 & " KB"
	  	if iFileSize<1000 then sDiskSize=iFileSize& " bytes"
	  	sUploadedBy="Uploaded by "&rsFiles("Staff_First_Name")&" "&rsFiles("Staff_Surname") & " on " &CvbShortdateTime(rsFiles("F_File_Uploaded"),true)
	  	if rsFiles("staff_sec_level")<iSec or session("Staff_ID")=rsFiles("F_FileUploaded_User") then bDelFile=1
	  	sIconName="attach.gif"
	  	if instr(f1.name,".")>0 then 
	  		sExt=Right(f1.name,Len(f1.name)-instr(f1.name,".")+1)
	  	else 
	  		sExt=""
	  	end if
	  	'response.Write(sExt)
	  	'response.end
	  	Select case sExt 
	  		case ".gif",".jpg",".ico",".png",".tiff"
	  			sIconName="img.gif"
	  		case ".pdf"
	  			sIconName="pdf.gif"
	  		case ".xls",".xlsx",".csv"
	  			sIconName="xls.gif"
	   		case ".pdf"
	  			sIconName="pdf.gif"
	   		case ".txt"
	  			sIconName="file.gif"
	  	end Select 
	  	'Get File ID
	  	iFiles=iFiles+1
	  	if iFiles mod 2=1 then sRowClass="trSeperate" else sRowClass=""
	    s=s&"<tr class="""&sRowClass&""" id=""fr"&iFileID&"""><td>"&iFiles&".</td> <td><img src=""/images/fileicons/"&sIconName&"""></td>"
	    s=s&" <td title="""&sUploadedBy&"""><a href="""&URLLink&""&f1.name&""">"&f1.name&" ("&sDiskSize&")</a></td>"
	    s=s&" <td><a class=""remove"" href=""#"" onclick=""rmfl('"&iFileID&"','"&f1.name&"')""> remove</a></td> "
	    s=s&"</tr>"
  	end if
    rsFiles.close
  Next
  ShowFileList = s&sErrMsg
  if ShowFileList<>"" then 
  	
  	sheader="<table><tr><td colspan=""4"" class=""tblfiles"">Files attached to this record and field</td></tr>"
  	ShowFileList=sheader&ShowFileList&"</table>"
  end if
  set rsFiles=nothing
  set f=nothing
  set fc=nothing
  Set fso = Nothing
End Function


function ConvertToHTML_InputBasic(sName,sValue,sFieldType,sRequired,iSize,sReadOnly,sInputClass,bShown,bFile,sKey)
select case sFieldType
case "bit"
	sType="Checkbox"
	intFieldLength=1
	intMaxFieldLength=1
	HTML_Value=" Value=""true"""
	if sValue=true or ucase(sValue)="TRUE" or sValue="1" then 
		sChecked=" checked"
	else
		sChecked=""
	end if
case "datetime"
	sType="Text"
	intMaxFieldLength=iSize
	intFieldLength=iSize+2
	HTML_Value=" Value='" & CvbShortdate(sValue) & "'"
case "nvarchar","nchar","varchar"
	sType="Text"
	if instr(ucase(sName),"PASSWORD")>0 and session("staff_sec_Level")<4 then sType="Password"
	intMaxFieldLength=iSize
	intFieldLength=50 + cint(int((iSize-50)*2/3))
	if intFieldLength>100 then intFieldLength=80
	if iSize>256 and bShown=true and session("rowView")=True then
		iRows=cint(iSize/80)
		if iRows>15 then iRows=15
		'ConvertToHTML_Input="<textarea cols=80 rows="&iRows&" name=""" & sName & """ " & sReadonly & ">" & sValue & "</textarea>" 
	end if
	if isnull(sValue) then sValue=""
	'sValue=replace(sValue,"""","''")
	if instr(sValue,"""")>0 and instr(sValue,"'")=0 then
		HTML_Value=" Value='" & sValue & "'"
	else
		if instr(sValue,"""")>0 then sValue=replace(sValue,"""","")
		HTML_Value=" Value=""" & sValue & """"
	end if
case else
	sType="Text"
	intMaxFieldLength=iSize+3
	intFieldLength=cint(iSize*1)
	if isnull(sValue) then sValue=""
	sValue=replace(sValue,"""","''")
	HTML_Value=" Value=""" & sValue & """"
end Select
if bShown=false then
	sType="Hidden"
end if
if sReadOnly=true  then sReadonly=" READONLY " else sReadonly=""
if sInputClass<>"" then sClass=" class=""" & sInputClass &""" "
if ConvertToHTML_InputBasic="" then ConvertToHTML_InputBasic="<input type=""" & sType & """ "& sClass & " name="""& sName &""" id="""& sName &""" size=""" & intFieldLength & """ maxlength=""" & intMaxFieldLength & """ " & HTML_Value & " " & sReadonly & sChecked &">"
end function


sub Get_Image(sURL,sSavePath)
'on error resume next
Set xmlhttp = CreateObject("Msxml2.SERVERXMLHTTP")
response.Write(sURL&"</br>")
xmlhttp.Open "GET", sURL, false
xmlhttp.Send()	
'Create a Stream
Set adodbStream = CreateObject("ADODB.Stream")
status = xmlhttp.status 
if err.number <> 0 or status <> 200 then 
    if status = 404 then 
        Response.Write "Page does not exist (404)." 
    elseif status >= 401 and status < 402 then 
        Response.Write "Access denied (401)." 
    elseif status >= 500 and status <= 600 then 
        Response.Write "500 Internal Server Error on remote site." 
    else 
        Response.write "Server is down or does not exist." 
    end if 
else  
    Response.Write "Server is up and URL is available."  
end if  
'Open the stream
adodbStream.Open 
adodbStream.Type = 1 'adTypeBinary
adodbStream.Write xmlhttp.responseBody
'response.end
x=rwb("writing to : "&sSavePath)
adodbStream.SaveToFile sSavePath, 2 'adSaveCreateOverWrite
response.write sSavePath&" saved!</br>"
adodbStream.Close
Set adodbStream = Nothing
Set xmlhttp = Nothing
'on error goto 0				
end sub

sub Resize_photos(sPath,sSourcePath,sPN,iID,sImgName)
on error resume next
set fs = CreateObject("Scripting.FileSystemObject")
set jpeg = Server.CreateObject("Persits.Jpeg")
Set Jpeg2 = Server.CreateObject("Persits.Jpeg")
dim arrSizes(10)
dim arrPaths(10)
arrSizes(0)=800
arrSizes(1)=800
arrSizes(2)=400
arrSizes(3)=200
arrSizes(4)=100
arrSizes(5)=50
arrPaths(0)="\x-large\"
arrPaths(1)="\large\"
arrPaths(2)="\normal\"
arrPaths(3)="\small\"
arrPaths(4)="\x-small\"
arrPaths(5)="\xx-small\"
'sPath="E:\Inetpub\wwwroot\Mother\Images\DB2\Stock"
sAddPhoto=""
'check path make sure photo exists...
if fs.FileExists(sSourcePath) then 
	Jpeg.open sSourcePath
	'response.write sSourcePath
	
	For i=0 to 5 'Different sizes
		on error resume next
		iWidth=arrSizes(i)
		Jpeg.Width = iWidth
		Jpeg.Height = Jpeg.OriginalHeight * iWidth / Jpeg.OriginalWidth
		'jpeg.Height = jpeg.OriginalHeight * Upload.Form("scale") / 100
		SavePath = sPath & arrPaths(i) & sImgName
		if not fs.FolderExists (sPath&arrPaths(i)) then
			fs.CreateFolder(sSavePath&arrPaths(i))
		end if 
		Jpeg.Save(SavePath)
	  on error goto 0
	  'Response.Write SavePath & " Saved!" & "</br>"
	Next
	on error goto 0
	call OpenDBA()
	connA.execute ("Update Stock set Is_visable=1,has_Photo=1 where ID=" & iID & " and DB_owner="& idb)
	call CloseDBA()
end if
set jpeg = nothing
Set Jpeg2 = nothing
on error goto 0
End sub


sub DownloadFile(sFullPath)
	Response.Buffer = True
	Response.Clear
	' set the directory that contains the files here
	strFileName = sFullPath
	sFiles=split(sFullPath,"\")
	sFname= sFiles(ubound(sFiles))
	Set Sys = Server.CreateObject( "Scripting.FileSystemObject" )
	Set Bin = Sys.OpenTextFile( strFileName, 1, False )
	If Sys.FileExists( strFileName ) Then
		' This is the importent part :-)
		'  Set the Filename to save as
		Call Response.AddHeader( "Content-Disposition", "attachment; filename=" &sFName )
		' Make sure the browser downloads, instead of running it
		Response.ContentType = "application/octet-stream"
		' Send as a Binary Byte Stream
		While Not Bin.AtEndOfStream
			Response.BinaryWrite( ChrB( Asc( Bin.Read( 1 ) ) ) )
		Wend
	Else
		Response.Redirect( "/home.asp" )
	End If
	Bin.Close : Set Bin = Nothing
	Set Sys = Nothing
end sub
%>
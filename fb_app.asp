<%

'' Classic ASP Facebook: Version 4.0 01/16/2012
'' by Larry Boeldt
''
''  This library provides the foundation for authenticating with Facebook. 
'' The goal is to provide a mechanism by which users can use Facebook to register
'' on your website rather than having to fill out a registration form.  
''
'' Populate your Facebook API Key in FACEBOOK_APP_ID and your Application Secret
'' in FACEBOOK_SECRET 
'' 
''
'' Optionally update the scope variable to gain more access to the user's profile
'' for more options to put in the in scope variable see: 
'' http://developers.facebook.com/docs/reference/api/permissions/
''
'' 
''
 


'' You no longer need to set this, every facebook auth page
'' should include fb_app (this file) thus always authenticate
'' this setting now ensures the user always ends up where they
'' started before the auth request
dim FACEBOOK_REDIR_URL: FACEBOOK_REDIR_URL = GetRedirURL()
session("FACEBOOK_REDIR_URL") = FACEBOOK_REDIR_URL


dim FACEBOOK_APP_ID: FACEBOOK_APP_ID = "1575727666027370"
dim FACEBOOK_SECRET: FACEBOOK_SECRET = "439fd18b7706fb32ae392195827f11af"
dim FACEBOOK_SCOPE: FACEBOOK_SCOPE = "email"


'' This is a dictionary object that contains the
'' variables in the facebook cookie
dim cookie 
set cookie = get_facebook_cookie()
            
            
dim str_facebook_url            
if request.querystring("state") = session("state") and request.querystring("state") <> "" then
     str_facebook_url = FACEBOOK_REDIR_URL 
     
     if session("FACEBOOK_REDIR_QS")<>"" then 
        str_facebook_url = str_facebook_url & "?" & session("FACEBOOK_REDIR_QS")
     end if
     
     session("FACEBOOK_REDIR_QS") = ""
     
     response.redirect str_facebook_url
end if

function get_facebook_cookie() 
	dim sCookie
	dim arCookie
	dim payload
	dim arItm
	dim sItm
	dim iX
	dim enc
	dim key

	dim app_id
	dim app_secret
	dim my_url
	dim dialog_url
	dim token_url
	dim resp
	dim token
	dim expires
	dim graph_url
	dim json_str
	dim user
	dim code
	dim bExpired 
    dim fbu

    
	app_id = FACEBOOK_APP_ID
	app_secret = FACEBOOK_SECRET
	my_url = FACEBOOK_REDIR_URL

    set fbu = new fb_utility
	set enc = new encrypt
	dim args: set args = createobject("scripting.dictionary")  
	args.CompareMode = 1

	'' Set a default value for token expires
    if not isDate( request.cookies("token_expires") & "" ) then
		response.cookies("token_expires") = now() - 365
	end if

	bExpired = false
	if CDate( request.cookies("token_expires") & "" ) < now() _
	and request.cookies("access_token")<>"" then
		bExpired = true
		response.cookies("access_token") = ""
		
	end if

	sCookie = request.cookies("access_token") & ""
	' URL Decode the cookie


    session("step")="COOKIE CHECK " & now()
  	if (sCookie & "" = "") then 
        session("step") = "GOT COOKIE"
        if session("FACEBOOK_REDIR_QS")="" then session("FACEBOOK_REDIR_QS")=request.querystring
        
	    '' oAuth 2.0 - 
	    code = request.querystring("code")

		'' Get authorization
	    if isempty(code) then
	        session("state") = enc.md5(fbu.newid()) '' CSRF protection
	        dialog_url = "http://www.facebook.com/dialog/oauth" & _
	            "?client_id=" & app_id & _
	            "&redirect_uri=" & server.urlencode(my_url) & _
	            "&state=" & session("state") & _
				"&scope=" & FACEBOOK_SCOPE

            SetLoginStatus "LOGGING_IN"
	        response.redirect( dialog_url )
	    end if

		
	    if request.querystring("state") = session("state") then
	        token_url = "https://graph.facebook.com/oauth/access_token" & _
	            "?client_id=" & app_id & _
	            "&redirect_uri=" & server.urlencode(my_url) & _
	            "&client_secret=" & app_secret & _
	            "&code=" & code 

   
	        resp = fbu.get_page_contents(token_url)
            
	        token = fbu.parse_access_token( resp )
	        expires = fbu.parse_expires( resp )

	        graph_url = "https://graph.facebook.com/me?access_token=" & token

	        json_str = fbu.get_page_contents( graph_url )

            '' Expire token if logged out 
            if (lcase(mid( json_str, 3, 5 ))) = "error" then
                response.cookies("access_token") = ""
                response.cookies("token_expires") = ""
                if session("FACEBOOK_REDIR_QS") = "" then 
                    response.redirect FACEBOOK_REDIR_URL
                else
                    response.redirect FACEBOOK_REDIR_URL & "?" & session("FACEBOOK_REDIR_QS")
                end if
            end if

            
			set user = JSON.parse( json_str )

            session("json") = json_str
            
			response.cookies("access_token") = token
			response.cookies("token_expires") = now()+0.04
			response.cookies("facebook_user_id") = user.id
			response.cookies("facebook_user_name") = user.name
            
            SetLoginStatus "LOGGED_IN"
		end if
    else
        session("step") = "DONT_GOT_COOKIE"
	end if

	args.add "token", request.cookies("access_token") & ""
	args.add "expires", request.cookies("token_expires") & ""
	args.add "uid", request.cookies("facebook_user_id") & ""
	args.add "name", request.cookies("facebook_user_name") & ""

	set get_facebook_cookie = args
end function

function GetFBLoginStatus()
    GetLoginStatus = session("FB_LOGIN_STATUS")
end function

function SetLoginStatus(str)
    session("FB_LOGIN_STATUS") = str
end function 

function GetRedirURL()
	dim sResult
	dim url
	dim http
	dim host

	url = request.servervariables("url")
	host = request.servervariables("server_name")


	'' now figure out if we are http or https
	if request.servervariables("server_port_secure") & "" = "1" then
		http = "https://"
	else
		http = "http://"
	end if


	'' now build the full url
	sResult = http & host & url

	GetRedirURL = sResult
end function


%>
    
<%

'' End page loading utility


%>



<%

class fb_utility
    dim dictKey 
    dim dictItem

    sub Class_Initialize()
        '' Initialize KSort values
        dictKey  = 1
        dictItem = 2
    end sub
    
    '***************************************************************************
    ' Copyright 2010 Larry Boeldt 
    ' Function: get_page_contents  V2.0    2010-07-14
    ' Author: Larry Boeldt - lboeldt@scs-inc.net; larry@boeldt.net
    ' Parameter(s):
    '   strXML		A valid XML String
    '	strURL		The URL to post the XML string to
    ' Return Value:
    '   The return string from the post operation
    '
    ' Description:
    '	Use the Microsoft.ServerXMLHTTP object to GET data from a web page
    '   I use version 6.0 in this code example you may need to change
    '   to version 4.0 on some servers using
    '   CreateObject("MSXML2.ServerXMLHTTP.4.0")
    '***************************************************************************
    public function get_page_contents(strURL)
        Dim objHTTP
        dim strData
        dim Q

        Q = chr(34)       
        
        Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        objHTTP.Open "GET", strURL, False
        objHTTP.Send 
        strData = objHTTP.ResponseText

        if objHTTP.status = 200 Then
            strData = objHTTP.ResponseText
        else
            '' We know we have an error, create a standard error json (id=-1)
            strData = "{" & _
                Q & "id" & Q & ":" & Q & "-1" & Q & _
                "," & Q & "data" & Q & ":" & objHTTP.ResponseText & "}"
                
            '' ok let's check for a user logged out message
            if instr(1, strData, "user logged out", 1) > 0 then 
                response.cookies("access_token")=""
                set cookie = get_facebook_cookie()
            end if
        end if        
        
        set objHTTP = nothing       
        
        get_page_contents = strData
    end function 
    


    '***************************************************************************
    ' Copyright 2010 Larry Boeldt 
    ' Function: get_page_contents  V1.0    2010-07-14
    ' Author: Larry Boeldt - larryboeldt@msn.com
    ' Parameter(s):
    '   strXML		A valid XML String
    '	strURL		The URL to post the XML string to
    ' Return Value:
    '   The return string from the post operation
    '
    ' Description:
    '	Use the Microsoft.ServerXMLHTTP object to GET data from a web page
    '   I use version 6.0 in this code example you may need to change
    '   to version 4.0 on some servers using
    '   CreateObject("MSXML2.ServerXMLHTTP.4.0")
    '***************************************************************************
    public function post_page(strURL, strData)
        Dim objHTTP
        dim strResponse
        dim Q

        Q = chr(34)

        Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        objHTTP.Open "POST", strURL
        objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        objHTTP.setRequestHeader "Content-Length", Len(strData)
        objHTTP.Send strData
        
        
        if objHTTP.status = 200 Then
            strResponse = objHTTP.ResponseText
        else
            strResponse = "{" & _
                Q & "id" & Q & ":" & Q & "-1" & Q & _
                "," & Q & "data" & Q & ":" & objHTTP.ResponseText & "}"
        end if

         
        set objHTTP = nothing

        post_page = strResponse
    end function 
    
    '***************************************************************************
    ' Copyright 2010 Larry Boeldt 
    ' Function: NewID  V1.0    2012-01-16
    ' Author: Larry Boeldt - larry@boeldt.net
    ' Parameter(s):
    '   none
    ' Return Value:
    '   A standard GUID
    '
    ' Description:
    '     Generate a GUID using the scriptlet type library
    '***************************************************************************
    public Function NewID()
        dim TypeLib
        dim tg
        
        Set TypeLib = CreateObject("Scriptlet.TypeLib")
        tg = TypeLib.Guid
        Set TypeLib = Nothing
        NewID = left(tg, len(tg)-2)
    End Function        



     ''
    '' Dictionary Sort
    '' This is from microsoft KB246067
    '' http://support.microsoft.com/kb/246067
    ''   
    ''
    '' use SortDictionary to provide a php like ksort
    ''
    public function ksort(byref objDict )
        SortDictionary objDict, dictKey 
    end function
    
    Function SortDictionary(byref objDict, intSort)
      ' declare our variables
      Dim strDict()
      Dim objKey
      Dim strKey,strItem
      Dim X,Y,Z
      dim ut
      

      ' get the dictionary count
      Z = objDict.Count

      ' we need more than one item to warrant sorting
      If Z > 1 Then
        ' create an array to store dictionary information
        ReDim strDict(Z,2)
        X = 0
        ' populate the string array
        For Each objKey In objDict
            strDict(X,dictKey)  = CStr(objKey)
            strDict(X,dictItem) = CStr(objDict(objKey))
            X = X + 1
        Next

        ' perform a a shell sort of the string array
        For X = 0 to (Z - 2)
          For Y = X to (Z - 1)
            If StrComp(strDict(X,intSort),strDict(Y,intSort),vbTextCompare) > 0 Then
                strKey  = strDict(X,dictKey)
                strItem = strDict(X,dictItem)
                strDict(X,dictKey)  = strDict(Y,dictKey)
                strDict(X,dictItem) = strDict(Y,dictItem)
                strDict(Y,dictKey)  = strKey
                strDict(Y,dictItem) = strItem
            End If
          Next
        Next

        ' erase the contents of the dictionary object
        objDict.RemoveAll

        ' repopulate the dictionary with the sorted information
        For X = 0 to (Z - 1)
          objDict.Add strDict(X,dictKey), strDict(X,dictItem)
        Next

      End If

    End Function    
    '' end dictionary sort    
    
    

    '***************************************************************************
    ' Copyright 2010 Larry Boeldt 
    ' Function: parse_access_token  V1.0    2010-07-14
    ' Author: Larry Boeldt - larry@boeldt.net
    ' Parameter(s):
    '   str		String of data to parse
    ' Return Value:
    '   The return string contains the facebook access token
    '
    ' Description:
    ' parse the value of "access_token" from a query string
    '***************************************************************************
    public function parse_access_token( str )
        dim sResult 
        dim lStart
        dim lEnd
        
        lStart = instr( 1, str, "access_token=", 1 )
        if lStart > 0 then 
            lStart = lStart + len("access_token=")
            lEnd = instr( lStart, str, "&" )
            if lEnd = 0 then lEnd = len( str )
            
            sResult = mid( str, lStart, lEnd - lStart )
        end if
        
        parse_access_token = sResult
    end function


    '***************************************************************************
    ' Copyright 2010 Larry Boeldt 
    ' Function: parse_expires  V1.0    2012-01-16
    ' Author: Larry Boeldt - larry@boeldt.net
    ' Parameter(s):
    '   str		String of data to parse
    ' Return Value:
    '   The return string contains the facebook access token expire code
    '
    ' Description:
    ' 	Parse the value "expires" from a query string
    '***************************************************************************
    public function parse_expires( str )
        dim sResult 
        dim lStart
        dim lEnd
        
        lStart = instr( 1, str, "expires=", 1 )
        if lStart > 0 then 
            lStart = lStart + len("expires=")
            lEnd = instr( lStart, str, "&" )
            if lEnd = 0 then lEnd = len( str )
            
            sResult = mid( str, lStart, lEnd - lStart )
        end if
        
        
        parse_expires = sResult
    end function
    
end class
%>


<% 
' Credit for this impressively clever and elegant 
' implementation of URLDecode goes to Flangy Development. 
' Code Source: http://flangy.com/dev/asp/urldecode.html

'
' The neat thing is the use of GetRef to get a function pointer 
' combined with a regular expression to call the replace function
' all looping and rescursion is implicit and making the code 
' incredibly clean although it is a little obscure 

' The downside of these functions is they cannot be wrapped
' in a class. However you can place them into a .wsc file
' and they will work fine. 

' An inverse to Server.URLEncode
public function URLDecode(str)
    dim re
    set re = new RegExp

    str = Replace(str, "+", " ")
    
    re.Pattern = "%([0-9a-fA-F]{2})"
    re.Global = True
    URLDecode = re.Replace(str, GetRef("URLDecodeHex"))
end function

' Replacement function for the above
public function URLDecodeHex(match, hex_digits, pos, source)
    URLDecodeHex = chr("&H" & hex_digits)
end function



' A shortcut for server.urlencode
public Function URLEncode(str) 
    URLEncode = Server.URLEncode(str) 
End Function 
%>








<%
class encrypt
    'md5 implementation @0-4FCF0748
    ' Derived from the RSA Data Security, Inc. MD5 Message-Digest Algorithm,
    ' as set out in the memo RFC1321.
    '
    '
    ' ASP VBScript code for generating an MD5 'digest' or 'signature' of a string. The
    ' MD5 algorithm is one of the industry standard methods for generating digital
    ' signatures. It is generically known as a digest, digital signature, one-way
    ' encryption, hash or checksum algorithm. A common use for MD5 is for password
    ' encryption as it is one-way in nature, that does not mean that your passwords
    ' are not free from a dictionary attack.
    '
    ' This is 'free' software with the following restrictions:
    '
    ' You may not redistribute this code as a 'sample' or 'demo'. However, you are free
    ' to use the source code in your own code, but you may not claim that you created
    ' the sample code. It is expressly forbidden to sell or profit from this source code
    ' other than by the knowledge gained or the enhanced value added by your own code.
    '
    ' Use of this software is also done so at your own risk. The code is supplied as
    ' is without warranty or guarantee of any kind.
    '
    ' Should you wish to commission some derivative work based on this code provided
    ' here, or any consultancy work, please do not hesitate to contact us.
    '
    ' Web Site:  http://www.frez.co.uk
    ' E-mail:    sales@frez.co.uk

    dim BITS_TO_A_BYTE 
    dim BYTES_TO_A_WORD 
    dim BITS_TO_A_WORD 

    Dim m_lOnBits(30)
    Dim m_l2Power(30)
     
    sub class_initialize() ' init
        BITS_TO_A_BYTE = 8
        BYTES_TO_A_WORD = 4
        BITS_TO_A_WORD = 32
        
        m_lOnBits(0) = CLng(1)
        m_lOnBits(1) = CLng(3)
        m_lOnBits(2) = CLng(7)
        m_lOnBits(3) = CLng(15)
        m_lOnBits(4) = CLng(31)
        m_lOnBits(5) = CLng(63)
        m_lOnBits(6) = CLng(127)
        m_lOnBits(7) = CLng(255)
        m_lOnBits(8) = CLng(511)
        m_lOnBits(9) = CLng(1023)
        m_lOnBits(10) = CLng(2047)
        m_lOnBits(11) = CLng(4095)
        m_lOnBits(12) = CLng(8191)
        m_lOnBits(13) = CLng(16383)
        m_lOnBits(14) = CLng(32767)
        m_lOnBits(15) = CLng(65535)
        m_lOnBits(16) = CLng(131071)
        m_lOnBits(17) = CLng(262143)
        m_lOnBits(18) = CLng(524287)
        m_lOnBits(19) = CLng(1048575)
        m_lOnBits(20) = CLng(2097151)
        m_lOnBits(21) = CLng(4194303)
        m_lOnBits(22) = CLng(8388607)
        m_lOnBits(23) = CLng(16777215)
        m_lOnBits(24) = CLng(33554431)
        m_lOnBits(25) = CLng(67108863)
        m_lOnBits(26) = CLng(134217727)
        m_lOnBits(27) = CLng(268435455)
        m_lOnBits(28) = CLng(536870911)
        m_lOnBits(29) = CLng(1073741823)
        m_lOnBits(30) = CLng(2147483647)
        
        m_l2Power(0) = CLng(1)
        m_l2Power(1) = CLng(2)
        m_l2Power(2) = CLng(4)
        m_l2Power(3) = CLng(8)
        m_l2Power(4) = CLng(16)
        m_l2Power(5) = CLng(32)
        m_l2Power(6) = CLng(64)
        m_l2Power(7) = CLng(128)
        m_l2Power(8) = CLng(256)
        m_l2Power(9) = CLng(512)
        m_l2Power(10) = CLng(1024)
        m_l2Power(11) = CLng(2048)
        m_l2Power(12) = CLng(4096)
        m_l2Power(13) = CLng(8192)
        m_l2Power(14) = CLng(16384)
        m_l2Power(15) = CLng(32768)
        m_l2Power(16) = CLng(65536)
        m_l2Power(17) = CLng(131072)
        m_l2Power(18) = CLng(262144)
        m_l2Power(19) = CLng(524288)
        m_l2Power(20) = CLng(1048576)
        m_l2Power(21) = CLng(2097152)
        m_l2Power(22) = CLng(4194304)
        m_l2Power(23) = CLng(8388608)
        m_l2Power(24) = CLng(16777216)
        m_l2Power(25) = CLng(33554432)
        m_l2Power(26) = CLng(67108864)
        m_l2Power(27) = CLng(134217728)
        m_l2Power(28) = CLng(268435456)
        m_l2Power(29) = CLng(536870912)
        m_l2Power(30) = CLng(1073741824)
    end sub

    Private Function LShift(lValue, iShiftBits)
        If iShiftBits = 0 Then
            LShift = lValue
            Exit Function
        ElseIf iShiftBits = 31 Then
            If lValue And 1 Then
                LShift = &H80000000
            Else
                LShift = 0
            End If
            Exit Function
        ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
            Err.Raise 6
        End If

        If (lValue And m_l2Power(31 - iShiftBits)) Then
            LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
        Else
            LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
        End If
    End Function

    Private Function RShift(lValue, iShiftBits)
        If iShiftBits = 0 Then
            RShift = lValue
            Exit Function
        ElseIf iShiftBits = 31 Then
            If lValue And &H80000000 Then
                RShift = 1
            Else
                RShift = 0
            End If
            Exit Function
        ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
            Err.Raise 6
        End If
        
        RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)

        If (lValue And &H80000000) Then
            RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
        End If
    End Function

    Private Function RotateLeft(lValue, iShiftBits)
        RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
    End Function

    Private Function AddUnsigned(lX, lY)
        Dim lX4
        Dim lY4
        Dim lX8
        Dim lY8
        Dim lResult
     
        lX8 = lX And &H80000000
        lY8 = lY And &H80000000
        lX4 = lX And &H40000000
        lY4 = lY And &H40000000
     
        lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
     
        If lX4 And lY4 Then
            lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
        ElseIf lX4 Or lY4 Then
            If lResult And &H40000000 Then
                lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
            Else
                lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
            End If
        Else
            lResult = lResult Xor lX8 Xor lY8
        End If
     
        AddUnsigned = lResult
    End Function

    Private Function F(x, y, z)
        F = (x And y) Or ((Not x) And z)
    End Function

    Private Function G(x, y, z)
        G = (x And z) Or (y And (Not z))
    End Function

    Private Function H(x, y, z)
        H = (x Xor y Xor z)
    End Function

    Private Function I(x, y, z)
        I = (y Xor (x Or (Not z)))
    End Function

    Private Sub FF(a, b, c, d, x, s, ac)
        a = AddUnsigned(a, AddUnsigned(AddUnsigned(F(b, c, d), x), ac))
        a = RotateLeft(a, s)
        a = AddUnsigned(a, b)
    End Sub

    Private Sub GG(a, b, c, d, x, s, ac)
        a = AddUnsigned(a, AddUnsigned(AddUnsigned(G(b, c, d), x), ac))
        a = RotateLeft(a, s)
        a = AddUnsigned(a, b)
    End Sub

    Private Sub HH(a, b, c, d, x, s, ac)
        a = AddUnsigned(a, AddUnsigned(AddUnsigned(H(b, c, d), x), ac))
        a = RotateLeft(a, s)
        a = AddUnsigned(a, b)
    End Sub

    Private Sub II(a, b, c, d, x, s, ac)
        a = AddUnsigned(a, AddUnsigned(AddUnsigned(I(b, c, d), x), ac))
        a = RotateLeft(a, s)
        a = AddUnsigned(a, b)
    End Sub

    Private Function ConvertToWordArray(sMessage)
        Dim lMessageLength
        Dim lNumberOfWords
        Dim lWordArray()
        Dim lBytePosition
        Dim lByteCount
        Dim lWordCount
        
        Const MODULUS_BITS = 512
        Const CONGRUENT_BITS = 448
        
        lMessageLength = Len(sMessage)
        
        lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
        ReDim lWordArray(lNumberOfWords - 1)
        
        lBytePosition = 0
        lByteCount = 0
        Do Until lByteCount >= lMessageLength
            lWordCount = lByteCount \ BYTES_TO_A_WORD
            lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
            lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
            lByteCount = lByteCount + 1
        Loop

        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE

        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)

        lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
        lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)
        
        ConvertToWordArray = lWordArray
    End Function

    Private Function WordToHex(lValue)
        Dim lByte
        Dim lCount
        
        For lCount = 0 To 3
            lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
            WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
        Next
    End Function

    Public Function MD5(sMessage)
        Dim x
        Dim k
        Dim AA
        Dim BB
        Dim CC
        Dim DD
        Dim a
        Dim b
        Dim c
        Dim d
        
        Const S11 = 7
        Const S12 = 12
        Const S13 = 17
        Const S14 = 22
        Const S21 = 5
        Const S22 = 9
        Const S23 = 14
        Const S24 = 20
        Const S31 = 4
        Const S32 = 11
        Const S33 = 16
        Const S34 = 23
        Const S41 = 6
        Const S42 = 10
        Const S43 = 15
        Const S44 = 21

        x = ConvertToWordArray(sMessage)
        
        a = &H67452301
        b = &HEFCDAB89
        c = &H98BADCFE
        d = &H10325476

        For k = 0 To UBound(x) Step 16
            AA = a
            BB = b
            CC = c
            DD = d
        
            FF a, b, c, d, x(k + 0), S11, &HD76AA478
            FF d, a, b, c, x(k + 1), S12, &HE8C7B756
            FF c, d, a, b, x(k + 2), S13, &H242070DB
            FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE
            FF a, b, c, d, x(k + 4), S11, &HF57C0FAF
            FF d, a, b, c, x(k + 5), S12, &H4787C62A
            FF c, d, a, b, x(k + 6), S13, &HA8304613
            FF b, c, d, a, x(k + 7), S14, &HFD469501
            FF a, b, c, d, x(k + 8), S11, &H698098D8
            FF d, a, b, c, x(k + 9), S12, &H8B44F7AF
            FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1
            FF b, c, d, a, x(k + 11), S14, &H895CD7BE
            FF a, b, c, d, x(k + 12), S11, &H6B901122
            FF d, a, b, c, x(k + 13), S12, &HFD987193
            FF c, d, a, b, x(k + 14), S13, &HA679438E
            FF b, c, d, a, x(k + 15), S14, &H49B40821
        
            GG a, b, c, d, x(k + 1), S21, &HF61E2562
            GG d, a, b, c, x(k + 6), S22, &HC040B340
            GG c, d, a, b, x(k + 11), S23, &H265E5A51
            GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA
            GG a, b, c, d, x(k + 5), S21, &HD62F105D
            GG d, a, b, c, x(k + 10), S22, &H2441453
            GG c, d, a, b, x(k + 15), S23, &HD8A1E681
            GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8
            GG a, b, c, d, x(k + 9), S21, &H21E1CDE6
            GG d, a, b, c, x(k + 14), S22, &HC33707D6
            GG c, d, a, b, x(k + 3), S23, &HF4D50D87
            GG b, c, d, a, x(k + 8), S24, &H455A14ED
            GG a, b, c, d, x(k + 13), S21, &HA9E3E905
            GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8
            GG c, d, a, b, x(k + 7), S23, &H676F02D9
            GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A
                
            HH a, b, c, d, x(k + 5), S31, &HFFFA3942
            HH d, a, b, c, x(k + 8), S32, &H8771F681
            HH c, d, a, b, x(k + 11), S33, &H6D9D6122
            HH b, c, d, a, x(k + 14), S34, &HFDE5380C
            HH a, b, c, d, x(k + 1), S31, &HA4BEEA44
            HH d, a, b, c, x(k + 4), S32, &H4BDECFA9
            HH c, d, a, b, x(k + 7), S33, &HF6BB4B60
            HH b, c, d, a, x(k + 10), S34, &HBEBFBC70
            HH a, b, c, d, x(k + 13), S31, &H289B7EC6
            HH d, a, b, c, x(k + 0), S32, &HEAA127FA
            HH c, d, a, b, x(k + 3), S33, &HD4EF3085
            HH b, c, d, a, x(k + 6), S34, &H4881D05
            HH a, b, c, d, x(k + 9), S31, &HD9D4D039
            HH d, a, b, c, x(k + 12), S32, &HE6DB99E5
            HH c, d, a, b, x(k + 15), S33, &H1FA27CF8
            HH b, c, d, a, x(k + 2), S34, &HC4AC5665
        
            II a, b, c, d, x(k + 0), S41, &HF4292244
            II d, a, b, c, x(k + 7), S42, &H432AFF97
            II c, d, a, b, x(k + 14), S43, &HAB9423A7
            II b, c, d, a, x(k + 5), S44, &HFC93A039
            II a, b, c, d, x(k + 12), S41, &H655B59C3
            II d, a, b, c, x(k + 3), S42, &H8F0CCC92
            II c, d, a, b, x(k + 10), S43, &HFFEFF47D
            II b, c, d, a, x(k + 1), S44, &H85845DD1
            II a, b, c, d, x(k + 8), S41, &H6FA87E4F
            II d, a, b, c, x(k + 15), S42, &HFE2CE6E0
            II c, d, a, b, x(k + 6), S43, &HA3014314
            II b, c, d, a, x(k + 13), S44, &H4E0811A1
            II a, b, c, d, x(k + 4), S41, &HF7537E82
            II d, a, b, c, x(k + 11), S42, &HBD3AF235
            II c, d, a, b, x(k + 2), S43, &H2AD7D2BB
            II b, c, d, a, x(k + 9), S44, &HEB86D391
        
            a = AddUnsigned(a, AA)
            b = AddUnsigned(b, BB)
            c = AddUnsigned(c, CC)
            d = AddUnsigned(d, DD)
        Next
        
        MD5 = LCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))
    End Function

    'End md5 implementation

end class

%>
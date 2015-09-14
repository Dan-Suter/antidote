<%

''
'' Facebook Graph API for Classic ASP V1.0 01/26/2012
'' Author: Larry Boeldt
'' 
''
'' Requires: fb_app.asp and json2.asp
''
'' REF: http://developers.facebook.com/docs/reference/api/
''
'' Graph API calls implemented
''  
'' UserInfo: https://graph.facebook.com/me?access_token=...
'' Profile feed (Wall): https://graph.facebook.com/me/feed?access_token=...
'' Friends: https://graph.facebook.com/me/friends?access_token=...
'' Photo Albums: https://graph.facebook.com/me/albums?access_token=...
'' 
'' 
'' Graph API calls not implemented (TODO LIST)
'' News feed: https://graph.facebook.com/me/home?access_token=...
'' Likes: https://graph.facebook.com/me/likes?access_token=...
'' Movies: https://graph.facebook.com/me/movies?access_token=...
'' Music: https://graph.facebook.com/me/music?access_token=...
'' Books: https://graph.facebook.com/me/books?access_token=...
'' Notes: https://graph.facebook.com/me/notes?access_token=...
'' Permissions: https://graph.facebook.com/me/permissions?access_token=...
'' Photo Tags: https://graph.facebook.com/me/photos?access_token=...
'' Video Tags: https://graph.facebook.com/me/videos?access_token=...
'' Video Uploads: https://graph.facebook.com/me/videos/uploaded?access_token=...
'' Events: https://graph.facebook.com/me/events?access_token=...
'' Groups: https://graph.facebook.com/me/groups?access_token=...
'' Checkins: https://graph.facebook.com/me/checkins?access_token=...
''
dim last_json_string
dim last_graph_url




class fb_graph_api


    ''**************************************************************************
    '' Copyright 2012 Larry Boeldt 
    '' Function: post_to_wall  V1.0    2010-07-14
    '' Author: Larry Boeldt - lboeldt@scs-inc.net; larry@boeldt.net
    '' Parameter(s):
    ''          strMessage 
    ''          strPicture
    ''          strLink
    ''          strName
    ''          strCaption
    ''          strDescription
    ''
    ''
    '' Return Value:
    ''   
    ''
    '' Description:
    ''    access_token – the access token of the account we are publishing for
    ''    message – the message to publish
    ''    picture (optional) – a link to the photo of the post
    ''    name (optional) – the “title” of the post
    ''    link (optional) – a link to where the name will go to when clicked
    ''    caption (optional) – a few words to describe the link/name
    ''    description (optional) – more than a few words to describe the link/name
    ''
    ''**************************************************************************
    function post_to_wall(strMessage, strPicture, strLink, strName, strCaption, strDescription)
        dim token
        dim url
        dim uid
        dim data
        dim sResult 
        dim obj 
        dim id
        dim fbu
        
        set fbu = new fb_utility
        
        uid = cookie("uid")
        token = cookie("token") & ""

        url = "https://graph.facebook.com/" & uid & "/feed"
        data = ""
        data = data & "access_token=" & token
        data = data & "&message=" & server.urlencode( strMessage )

        if strPicture <> "" then
            data = data & "&picture=" & strPicture 
        end if

        if strLink <> "" then 
            data = data & "&link=" & strLink 
        end if

        if strName <> "" then
            data = data & "&name=" & strName 
        end if
        
        if strCaption <> "" then
            data = data & "&caption=" & strCaption 
        end if
        
        if strDescription <> "" then
            data = data & "&description=" & strDescription 
        end if
        
         
        sResult = fbu.post_page( url, data )

        last_json_string = sResult
        
        set obj = JSON.parse( sResult )

        id = obj.id & ""

        if obj.id = "-1" then 
            response.write "ERROR!"
            response.write "<br/>"
            response.write obj.data.error.message
            response.write "<br/>"
        else
            response.write "SUCCESS!"
            response.write "<br/>"
        end if
        
        response.write "ID: " & id
        response.write "<br/>"
        response.write "Full Result String:<br/> "
        response.write sResult
        
    end function

    ''**************************************************************************
    '' Copyright 2012 Larry Boeldt 
    '' Function: get_user  V1.0    2012-01-26
    '' Author: Larry Boeldt - lboeldt@scs-inc.net; larry@boeldt.net
    '' Parameter(s):
    ''   None
    ''
    '' Return Value:
    ''   User object from parsed json string
    ''
    '' Description:
    ''    Call the graph API and get user object as parsed from json. 
    ''**************************************************************************
    function get_user()

        dim token
        dim uid
        dim fbu
        dim graph_url
        dim json_str
        
        set fbu = new fb_utility
        
        uid = cookie("uid")
        token = cookie("token") & ""
        
        graph_url = "https://graph.facebook.com/me?access_token=" & token

        json_str = fbu.get_page_contents( graph_url )

        last_json_string = json_str
        
        set user = JSON.parse( json_str )

        
        set get_user_info = user
    end function

    ''**************************************************************************
    '' Copyright 2012 Larry Boeldt 
    '' Function: get_object_JSON  V1.0    2012-02-10
    '' Author: Larry Boeldt - lboeldt@scs-inc.net; larry@boeldt.net
    '' Parameter(s):
    ''   None
    ''
    '' Return Value:
    ''   User object from parsed json string
    ''
    '' Description:
    ''    Call the graph API and get user object as parsed from json. 
    ''**************************************************************************
    function get_object_JSON(ID, DataSet)
        dim token
        dim uid
        dim fbu
        dim graph_url
        dim json_str
        
        set fbu = new fb_utility
        
        uid = cookie("uid")
        token = cookie("token") & ""
        
        
        graph_url = "https://graph.facebook.com/" & id & "?access_token=" & token

        json_str = fbu.get_page_contents( graph_url )

        last_json_string = json_str
        
        set get_object_JSON = json_str
    end function


    ''**************************************************************************
    '' Copyright 2012 Larry Boeldt 
    '' Function: get_friends  V1.0    2012-01-26
    '' Author: Larry Boeldt - lboeldt@scs-inc.net; larry@boeldt.net
    '' Parameter(s):
    ''     None
    ''
    '' Return Value:
    ''    JSON friends list with id and name fields
    ''
    '' Description:
    ''  Call graph api to get the current user's friend list
    ''**************************************************************************
    function get_friends()
        dim token
        dim uid
        dim graph_url 
        dim fbu
        dim json_str
        dim obj
        
        set fbu = new fb_utility
        
        uid = cookie("uid")
        token = cookie("token") & ""
        graph_url = "https://graph.facebook.com/me/friends?access_token=" & token

        last_graph_url = graph_url
        
        json_str = fbu.get_page_contents( graph_url )

        last_json_string = json_str

        set obj = JSON.parse( json_str )
       
        set get_friends = obj
    end function


    ''**************************************************************************
    '' Copyright 2012 Larry Boeldt 
    '' Function: get_friends  V1.0    2012-01-26
    '' Author: Larry Boeldt - lboeldt@scs-inc.net; larry@boeldt.net
    '' Parameter(s):
    ''     None
    ''
    '' Return Value:
    ''    JSON friends list with id and name fields
    ''
    '' Description:
    ''  Call graph api to get a friend of the current user
    ''**************************************************************************
    function get_friend(id)
        dim token
        dim uid
        dim graph_url 
        dim fbu
        dim json_str
        dim obj
        dim user
        
        set fbu = new fb_utility
        set user = new fb_user
        
        uid = cookie("uid")
        token = cookie("token") & ""
        graph_url = "https://graph.facebook.com/" & id & "?access_token=" & token

        last_graph_url = graph_url
        
        json_str = fbu.get_page_contents( graph_url )

        last_json_string = json_str

        set obj = JSON.parse( json_str )

        user.load_from_json_user obj
        
        set get_friend = user
    end function


    ''**************************************************************************
    '' Copyright 2012 Larry Boeldt 
    '' Function: likes  V1.0    2012-01-26
    '' Author: Larry Boeldt - lboeldt@scs-inc.net; larry@boeldt.net
    '' Parameter(s):
    ''     ID - The object ID you want the user to like
    ''
    '' Return Value:
    ''    Response
    ''
    '' Description:
    ''  Call graph api to like object id passed 
    ''**************************************************************************
    function likes(id)
        dim token
        dim uid
        dim graph_url 
        dim fbu
        dim json_str
        dim obj
        dim user
        
        set fbu = new fb_utility
        
        uid = cookie("uid")
        token = cookie("token") & ""
        graph_url = "https://graph.facebook.com/" & id & "/likes?access_token=" & token

        last_graph_url = graph_url
        
        json_str = fbu.post_page( graph_url, "" )

        last_json_string = json_str

        set obj = JSON.parse( json_str )

        likes = json_str
    end function
    '' Flowerbox --

end class

''**************************************************************************
'' Copyright 2012 Larry Boeldt 
'' Function: get_page_contents  V1.0    2012-01-26
'' Author: Larry Boeldt - lboeldt@scs-inc.net; larry@boeldt.net
'' Parameter(s):
''
''
'' Return Value:
''   
''
'' Description:
''**************************************************************************

%>



<%

''**************************************************************************
'' Copyright 2012 Larry Boeldt 
'' Class: facebook_user  V1.0    2012-01-26
'' Author: Larry Boeldt - lboeldt@scs-inc.net; larry@boeldt.net
''
'' Description:
''   Facebook user object
''
''**************************************************************************
class fb_user
    dim m_id '' The user's Facebook ID
    dim m_name '' The user's full name
    dim m_first_name '' The user's first name
    dim m_middle_name '' The user's middle name
    dim m_last_name '' The user's last name
    dim m_gender '' The user's gender: female or male
    dim m_locale '' The user's locale
    dim m_languages '' The user's languages
    dim m_link '' The URL of the profile for the user on Facebook
    dim m_username '' The user's Facebook username
    dim m_third_party_id '' An anonymous, but unique identifier for the user; only returned if specifically requested via the fields URL parameter
    dim m_timezone '' The user's timezone offset from UTC
    dim m_updated_time '' The last time the user's profile was updated; changes to the languages, link, timezone, verified, interested_in, favorite_athletes, favorite_teams, and video_upload_limits are not not reflected in this value
    dim m_verified '' The user's account verification status, either true or false (see below)
    dim m_bio '' The user's biography
    dim m_birthday '' The user's birthday
    dim m_education '' A list of the user's education history
    dim m_email '' The proxied or contact email address granted by the user
    dim m_hometown '' The user's hometown
    dim m_interested_in '' The genders the user is interested in
    dim m_location '' The user's current city
    dim m_political '' The user's political view
    dim m_favorite_athletes '' The user's favorite athletes; this field is deprecated and will be removed in the near future
    dim m_favorite_teams '' The user's favorite teams; this field is deprecated and will be removed in the near future
    dim m_quotes '' The user's favorite quotes
    dim m_relationship_status '' The user's relationship status: Single, In a relationship, Engaged, Married, It's complicated, In an open relationship, Widowed, Separated, Divorced, In a civil union, In a domestic partnership
    dim m_religion '' The user's religion
    dim m_significant_other '' The user's significant other
    dim m_video_upload_limits '' The size of the video file and the length of the video that a user can upload; only returned if specifically requested via the fields URL parameter
    dim m_website '' The URL of the user's personal website
    dim m_token ''
    dim m_json_str ''
    dim m_graph_url ''
    dim fbu '' Facebook utility object
    
    sub Class_Initialize()
        m_id = ""
        m_name = ""
        m_first_name = ""
        m_middle_name = ""
        m_last_name = ""
        m_gender = ""
        m_locale = ""
        m_languages = ""
        m_link = ""
        m_username = ""
        m_third_party_id = ""
        m_timezone = ""
        m_updated_time = ""
        m_verified = ""
        m_bio = ""
        m_birthday = ""
        m_education = ""
        m_email = ""
        m_hometown = ""
        m_interested_in = ""
        m_location = ""
        m_political = ""
        m_quotes = ""
        m_relationship_status = ""
        m_religion = ""
        m_significant_other = ""
        m_video_upload_limits = ""
        m_website = ""
        
        set fbu = new fb_utility
    end sub

    public property get graph_url()
        graph_url = m_graph_url
    end property
    
    public property get json_string()
        json_string = m_json_str
    end property
    
    public property get token()
        token = m_token
    end property
    public property let token(value)
        m_token = value
    end property
    
    public property get id()
        id = m_id
    end property
    public property let id(value) 
        m_id = value 
    end property

    public property get name()
        name = m_name
    end property
    public property let name(value) 
        m_name = value 
    end property

    public property get first_name()
        first_name = m_first_name
    end property
    public property let first_name(value) 
        m_first_name = value 
    end property

    public property get middle_name()
        middle_name = m_middle_name
    end property
    public property let middle_name(value) 
        m_middle_name = value 
    end property

    public property get last_name()
        last_name = m_last_name
    end property
    public property let last_name(value) 
        m_last_name = value 
    end property

    public property get gender()
        gender = m_gender
    end property
    public property let gender(value) 
        m_gender = value 
    end property

    public property get locale()
        locale = m_locale
    end property
    public property let locale(value) 
        m_locale = value 
    end property

    public property get languages()
        languages = m_languages
    end property
    public property let languages(value) 
        m_languages = value 
    end property

    public property get link()
        link = m_link
    end property
    public property let link(value) 
        m_link = value 
    end property

    public property get username()
        username = m_username
    end property
    public property let username(value) 
        m_username = value 
    end property

    public property get third_party_id()
        third_party_id = m_third_party_id
    end property
    public property let third_party_id(value) 
        m_third_party_id = value 
    end property

    public property get timezone()
        timezone = m_timezone
    end property
    public property let timezone(value) 
        m_timezone = value 
    end property

    public property get updated_time()
        updated_time = m_updated_time
    end property
    public property let updated_time(value) 
        m_updated_time = value 
    end property

    public property get verified()
        verified = m_verified
    end property
    public property let verified(value) 
        m_verified = value 
    end property

    public property get bio()
        bio = m_bio
    end property
    public property let bio(value) 
        m_bio = value 
    end property

    public property get birthday()
        birthday = m_birthday
    end property
    public property let birthday(value) 
        m_birthday = value 
    end property

    public property get education()
        education = m_education
    end property
    public property let education(value) 
        m_education = value 
    end property

    public property get email()
        email = m_email
    end property
    public property let email(value) 
        m_email = value 
    end property

    public property get hometown()
        hometown = m_hometown
    end property
    public property let hometown(value) 
        m_hometown = value 
    end property

    public property get interested_in()
        interested_in = m_interested_in
    end property
    public property let interested_in(value) 
        m_interested_in = value 
    end property

    public property get location()
        location = m_location
    end property
    public property let location(value) 
        m_location = value 
    end property

    public property get political()
        political = m_political
    end property
    public property let political(value) 
        m_political = value 
    end property

    public property get quotes()
        quotes = m_quotes
    end property
    public property let quotes(value) 
        m_quotes = value 
    end property

    public property get relationship_status()
        relationship_status = m_relationship_status
    end property
    public property let relationship_status(value) 
        m_relationship_status = value 
    end property

    public property get religion()
        religion = m_religion
    end property
    public property let religion(value) 
        m_religion = value 
    end property

    public property get significant_other()
        significant_other = m_significant_other
    end property
    public property let significant_other(value) 
        m_significant_other = value 
    end property

    public property get video_upload_limits()
        video_upload_limits = m_video_upload_limits
    end property
    public property let video_upload_limits(value) 
        m_video_upload_limits = value 
    end property

    public property get website()
        website = m_website
    end property
    public property let website(value) 
        m_website = value 
    end property

    public sub LoadMe()
        dim strFields
        dim obj
        

        
        m_graph_url = "https://graph.facebook.com/me?access_token=" & _
            m_token 
        
        m_json_str = fbu.get_page_contents( m_graph_url )
        
        set obj = JSON.parse( m_json_str )
        
        load_from_json_user obj
        
    end sub

    public sub Load(ID)
        dim strFields
        dim obj
        
        m_graph_url = "https://graph.facebook.com/" & ID & "?access_token=" & _
            m_token 
        
        m_json_str = fbu.get_page_contents( m_graph_url )
        
        set obj = JSON.parse( m_json_str )
        
        load_from_json_user obj
        
    end sub
    
    ''**************************************************************************
    '' Copyright 2012 Larry Boeldt 
    '' Function: load_from_json_user  V1.0    2012-01-26
    '' Author: Larry Boeldt - lboeldt@scs-inc.net; larry@boeldt.net
    '' Parameter(s):
    ''   fb_obj - The facebook object as returned from JSON.parse
    ''
    '' Description:
    ''   This function attempts to safely assign all user parameters from a 
    '' fasebook object so that you can consistently address the properties in 
    '' your code without concern about errors resulting from properties that
    '' are missing in the return json object
    ''**************************************************************************
    public sub load_from_json_user(fb_obj)
        on error resume next
        m_id = fb_obj.id
        m_name = fb_obj.name
        m_first_name = fb_obj.first_name
        m_middle_name = fb_obj.middle_name
        m_last_name = fb_obj.last_name
        m_gender = fb_obj.gender
        m_locale = fb_obj.locale
        m_languages = fb_obj.languages
        m_link = fb_obj.link
        m_username = fb_obj.username
        m_third_party_id = fb_obj.third_party_id
        m_timezone = fb_obj.timezone
        m_updated_time = fb_obj.updated_time
        m_verified = fb_obj.verified
        m_bio = fb_obj.bio
        m_birthday = fb_obj.birthday
        m_education = fb_obj.education
        m_email = fb_obj.email
        m_hometown = fb_obj.hometown.name
        m_interested_in = fb_obj.interested_in
        m_location = fb_obj.location.name
        m_political = fb_obj.political
        m_quotes = fb_obj.quotes
        m_relationship_status = fb_obj.relationship_status
        m_religion = fb_obj.religion
        m_significant_other = fb_obj.significant_other
        m_video_upload_limits = fb_obj.video_upload_limits
        m_website = fb_obj.website
        on error goto 0
        err.clear
    end sub
    

    
end class



%>


<%



class fb_photo

    '' **
    '' ** Declare Variables
    '' **
    
    '' Default Variables for properties 
    dim fbu '' Facebook utility object
    dim m_token ''
    dim m_json_str ''
    dim m_graph_url ''
    
    '' Default Internal Variables
    dim m_id '' unique id for this photo
    dim m_from '' Facebook From Object
    dim m_picture '' Picture URL
    dim m_source '' Source of the photo
    dim m_height '' Picture Height
    dim m_width '' Picture Width
    dim m_images '' list of images in varying sizes
    dim m_link '' Link to the Facebook image
    dim m_icon '' URL to the photos GIF icon
    dim m_created_time '' time the photo was created
    dim m_position '' position
    dim m_updated_time '' time the photo was updated
    dim m_comments '' comments regarding this photo
    dim m_likes '' likes of this photo
    
    
    '' Object Type Specific Variables
    
    '' **   
    '' ** property exposure
    '' **

    
    '' Default Properties
    public property get graph_url()
        graph_url = m_graph_url
    end property
    
    public property get json_string()
        json_string = m_json_str
    end property
    
    public property get token()
        token = m_token
    end property
    public property let token(value)
        m_token = value
    end property
    
    
    '' Object Type property exposure
    
    
    '' initialization
    sub Class_Initialize
        set fbu = new fb_utility
        m_id = ""
        m_from = ""
        m_picture = ""
        m_source = ""
        m_height = ""
        m_width = ""
        m_images = ""
        m_link = ""
        m_icon = ""
        m_created_time = ""
        m_position = ""
        m_updated_time = ""
        set m_comments = new List
        set m_likes = new List

    end sub
    
    '' termination
    sub Class_Terminate
        set fbu = nothing
        
    end sub

    '' **
    '' Implementation
    '' **
    public property get id()
        id = m_id
    end property
    public property let id(value) 
        m_id = value 
    end property

    public property get from()
        from = m_from
    end property
    public property let from(value) 
        m_from = value 
    end property

    public property get picture()
        picture = m_picture
    end property
    public property let picture(value) 
        m_picture = value 
    end property

    public property get source()
        source = m_source
    end property
    public property let source(value) 
        m_source = value 
    end property

    public property get height()
        height = m_height
    end property
    public property let height(value) 
        m_height = value 
    end property

    public property get width()
        width = m_width
    end property
    public property let width(value) 
        m_width = value 
    end property

    public property get images()
        images = m_images
    end property
    public property let images(value) 
        m_images = value 
    end property

    public property get link()
        link = m_link
    end property
    public property let link(value) 
        m_link = value 
    end property

    public property get icon()
        icon = m_icon
    end property
    public property let icon(value) 
        m_icon = value 
    end property

    public property get created_time()
        created_time = m_created_time
    end property
    public property let created_time(value) 
        m_created_time = value 
    end property

    public property get position()
        position = m_position
    end property
    public property let position(value) 
        m_position = value 
    end property

    public property get updated_time()
        updated_time = m_updated_time
    end property
    public property let updated_time(value) 
        m_updated_time = value 
    end property

    public property get comments_count
        comments_count = m_comments.count
    end property

    public property get comments(item)
        if isObject(m_comments(item)) then 
            set comments = m_comments(item)
        else 
            comments = m_comments(item)
        end if
    end property
    
    
    public property get likes_count
        likes_count = m_likes.count
    end property 
    
    public property get likes(item)
        if isObject( m_likes(item) ) then
            set likes = m_likes(item)
        else
            likes = m_likes(item)
        end if
    end property

    
    '' Load object by id
    public sub Load(ID)
        dim obj
        
        m_graph_url = "https://graph.facebook.com/" & id & "?access_token=" & _
            m_token 

        m_json_str = fbu.get_page_contents( m_graph_url )

        set obj = JSON.parse( m_json_str )
        
        load_from_json_object obj
        
    end sub
    
    public sub load_from_json_object( fb_obj )
        dim ix
        dim like1
        dim comment
        ''response.Write fb_obj.id & ":"
        on error resume next
        m_id = fb_obj.id
        m_from = fb_obj.from
        m_picture = fb_obj.picture
        m_source = fb_obj.source
        m_height = fb_obj.height
        m_width = fb_obj.width
        m_images = fb_obj.images
        m_link = fb_obj.link
        m_icon = fb_obj.icon
        m_created_time = fb_obj.created_time
        m_position = fb_obj.position
        m_updated_time = fb_obj.updated_time

        for ix = 0 to fb_obj.likes.data.length - 1
            set like1 = new fb_like
            like1.id = fb_obj.likes.data.get( ix ).id
            like1.name = fb_obj.likes.data.get( ix ).name
            m_likes.Add( like1 )
        next
                
        for ix = 0 to fb_obj.comments.data.length - 1
            set comment = new fb_comment
            comment.id = fb_obj.comments.data.get(ix).id
            comment.from.id = fb_obj.comments.data.get(ix).from.id
            comment.from.name = fb_obj.comments.data.get(ix).from.name
            comment.message = fb_obj.comments.data.get(ix).message
            comment.created_time = fb_obj.comments.data.get(ix).created_time
            m_comments.Add(comment)
            set comment = nothing
        next        

        on error goto 0
        err.clear
    end sub
    
    
end class '' fb_photo

class fb_album
    
    '' Declare Variables
    dim m_id '' facebook id of album
    dim m_from '' the facebook from object that is the source of this album
    dim m_name '' the name of the album
    dim m_description '' album description
    dim m_location '' name of location
    dim m_link '' link to facebook page
    dim m_cover_photo '' cover photo
    dim m_count '' number of pictures in the album
    dim m_type '' the type of album (normal)
    dim m_created_time '' the time the album was created
    dim m_updated_time '' the time the album was updated
    dim m_can_upload '' true if user can upload false if not
    dim m_likes '' an array of like objects
    dim m_comments '' an array of comment objects
    dim m_paging '' gets next page of 
    dim m_comments_count '' 
    dim m_likes_count '' 
    dim m_comments_size
    dim m_likes_size
    
    '' Default Variables for properties 
    dim m_token ''
    dim m_json_str ''
    dim m_graph_url ''
    
    '' Default Internal Variables
    
    '' Object Type Specific Variables
    
    '' **   
    '' ** property exposure
    '' **

    
    

    '' initialization
    sub Class_Initialize
        m_id = ""
        m_from = ""
        m_name = ""
        m_description = ""
        m_location = ""
        m_link = ""
        m_cover_photo = ""
        m_count = ""
        m_type = ""
        m_created_time = ""
        m_updated_time = ""
        m_can_upload = ""
        set m_likes = new List
        set m_comments = new List
        m_paging = ""
        m_comments_count = 0
        m_likes_count = 0
    end sub
    
    '' Default Properties
    public property get graph_url()
        graph_url = m_graph_url
    end property
    
    public property get json_string()
        json_string = m_json_str
    end property
    
    public property get token()
        token = m_token
    end property
    public property let token(value)
        m_token = value
    end property    
    
    public property get id()
        id = m_id
    end property
    public property let id(value) 
        m_id = value 
    end property
    
    public property get from()
        from = m_from
    end property
    public property let from(value) 
        m_from = value 
    end property
    
    public property get name()
        name = m_name
    end property
    public property let name(value) 
        m_name = value 
    end property
    
    public property get description()
        description = m_description
    end property
    public property let description(value) 
        m_description = value 
    end property
    
    public property get location()
        location = m_location
    end property
    public property let location(value) 
        m_location = value 
    end property
    
    public property get link()
        link = m_link
    end property
    public property let link(value) 
        m_link = value 
    end property
    
    public property get cover_photo()
        cover_photo = m_cover_photo
    end property
    public property let cover_photo(value) 
        m_cover_photo = value 
    end property
    
    public property get count()
        count = m_count
    end property
    public property let count(value) 
        m_count = value 
    end property
    
    public property get album_type()
        album_type = m_type
    end property
    public property let album_type(value) 
        m_type = value 
    end property
    
    public property get created_time()
        created_time = m_created_time
    end property
    public property let created_time(value) 
        m_created_time = value 
    end property
    
    public property get updated_time()
        updated_time = m_updated_time
    end property
    public property let updated_time(value) 
        m_updated_time = value 
    end property
    
    public property get can_upload()
        can_upload = m_can_upload
    end property
    public property let can_upload(value) 
        m_can_upload = value 
    end property
    
    public property get likes(item)
        if isObject( m_likes(item) ) then
            set likes = m_likes(item)
        else
            likes = m_likes( item )
        end if
    end property
    
    public property get comments(item)
        if isObject( m_comments(item) ) then 
            set comments = m_comments( item )
        else 
            comments = m_comments( item )
        end if
    end property


    public property get paging()
        paging = m_paging
    end property
    public property let paging(value) 
        m_paging = value 
    end property
    
    public property get comments_count()
        comments_count = m_comments.count
    end property
    
    public property get likes_count()
        likes_count = m_likes.count
    end property
    
    
    
    '' Implementation
    public sub Load(ID)
        dim obj
        dim fbu
        
        set fbu = new fb_utility
        
        m_graph_url = "https://graph.facebook.com/" & id & "?access_token=" & _
            m_token 

        m_json_str = fbu.get_page_contents( m_graph_url )
        
        
        
        set obj = JSON.parse( m_json_str )
        
        load_from_json_object obj        
    end sub
    
    public sub load_from_json_object( fb_obj )
        dim ix
        dim comment
        dim like1
        
        on error resume next
        m_id = fb_obj.id
        m_from = fb_obj.from.name
        m_name = fb_obj.name
        m_description = fb_obj.description
        m_location = fb_obj.location
        m_link = fb_obj.link
        m_cover_photo = fb_obj.cover_photo
        m_count = fb_obj.count
        m_type = fb_obj.type
        m_created_time = fb_obj.created_time
        m_updated_time = fb_obj.updated_time
        m_can_upload = fb_obj.can_upload
        
        for ix = 0 to fb_obj.likes.data.length - 1
            set like1 = new fb_like
            like1.id = fb_obj.likes.data.get( ix ).id
            like1.name = fb_obj.likes.data.get( ix ).name
            m_likes.Add( like1 )
        next
        
        for ix = 0 to fb_obj.comments.data.length - 1
            set comment = new fb_comment
            comment.id = fb_obj.comments.data.get(ix).id
            comment.from.id = fb_obj.comments.data.get(ix).from.id
            comment.from.name = fb_obj.comments.data.get(ix).from.name
            comment.message = fb_obj.comments.data.get(ix).message
            comment.created_time = fb_obj.comments.data.get(ix).created_time
            m_comments.Add(comment)
            set comment = nothing
        next
        on error goto 0        

        
        err.clear
        
    end sub
    
    
end class




class fb_from
    '' Declare Variables
    dim m_name '' name of source object
    dim m_category '' name of category of object
    dim m_id '' unique id of object


    sub Class_Initialize
        m_name = ""
        m_category = ""
        m_id = ""
    end sub
    
    public property get name()
        name = m_name
    end property
    public property let name(value) 
        m_name = value 
    end property
    
    public property get category()
        category = m_category
    end property
    public property let category(value) 
        m_category = value 
    end property
    
    public property get id()
        id = m_id
    end property
    public property let id(value) 
        m_id = value 
    end property
end class


class fb_like

    '' **
    '' ** Declare Variables
    '' **
   
    '' Default Variables for properties 
    dim fbu '' Facebook utility object
    
    '' Default Internal Variables
    dim m_id '' facebook id of person who likes
    dim m_name '' name of person who likes
    
    '' Object Type Specific Variables
    
    '' **   
    '' ** property exposure
    '' **
    
    '' Object Type property exposure
    public property get id()
        id = m_id
    end property
    public property let id(value) 
        m_id = value 
    end property

    public property get name()
        name = m_name
    end property
    public property let name(value) 
        m_name = value 
    end property

    
    
    '' initialization
    sub Class_Initialize
        m_id = ""
        m_name = ""
    end sub
    
    '' termination
    sub Class_Terminate
    end sub
    
end class


'' A Facebook DTO
class fb_comment

    '' **
    '' ** Declare Variables
    '' **
    
    '' Default Variables for properties 
    dim m_token ''
    dim m_json_str ''
    dim m_graph_url ''

    dim m_id '' facebook id of object
    dim m_message '' text of the message
    dim m_created_time '' time the comment was created
    dim m_likes '' The number of times this comment was liked
    dim m_user_likes '' This field is returned only if the authenticated user likes this comment
    dim m_type '' The type of this object; always returns comment
    dim m_from '' Name and id of person who posted this message
    
    '' Default Internal Variables
    
    '' Object Type Specific Variables
    
    '' **   
    '' ** property exposure
    '' **

    
    '' Default Properties
    public property get graph_url()
        graph_url = m_graph_url
    end property
    
    public property get json_string()
        json_string = m_json_str
    end property
    
    public property get token()
        token = m_token
    end property
    public property let token(value)
        m_token = value
    end property
    
    
    '' Object Type property exposure
    
    
    '' initialization
    sub Class_Initialize

        set m_from = new fb_from
        m_id = ""
        m_message = ""
        m_created_time = ""
        m_likes = ""
        m_user_likes = ""
        m_type = ""
        
    end sub
    
    '' termination
    sub Class_Terminate
        
    end sub

    '' **
    '' Implementation
    '' **
    public property get id()
        id = m_id
    end property
    public property let id(value) 
        m_id = value 
    end property
    
    public property get from_id()
        from_id = from.id
    end property
''    public property let from_id(value) 
''        m_from_id = value 
''    end property
    
    public property get from_name()
        from_name = from.name
    end property
''    public property let from_name(value) 
''        m_from_name = value 
''    end property


    public property get from()
        set from = m_from
    end property
    public property let from(value)
        if isObject( value ) then
            set m_from = value
        end if
    end property

    public property get message()
        message = m_message
    end property
    public property let message(value) 
        m_message = value 
    end property
    
    public property get created_time()
        created_time = m_created_time
    end property
    public property let created_time(value) 
        m_created_time = value 
    end property
    
    public property get likes()
        likes = m_likes
    end property
    public property let likes(value) 
        m_likes = value 
    end property
    
    public property get user_likes()
        user_likes = m_user_likes
    end property
    public property let user_likes(value) 
        m_user_likes = value 
    end property
    
    public property get TypeName()
        TypeName = m_type
    end property
    public property let TypeName(value) 
        m_type = value 
    end property
    
 
    

    
    '' Load object by id
    public sub Load(ID)
        dim obj
        dim fbu
        
        set fbu = new fb_utility
        
        m_graph_url = "https://graph.facebook.com/" & id & "?access_token=" & _
            m_token 

        m_json_str = fbu.get_page_contents( m_graph_url )
        
        
        
        set obj = JSON.parse( m_json_str )
        
        load_from_json_object obj
        
        set fbu = nothing
    end sub
    
    public sub load_from_json_object( fb_obj )
    
        on error resume next
        m_id = fb_obj.id
        m_from.id = fb_obj.from.id
        m_from.name = fb_obj.from.name
        m_message = fb_obj.message
        m_created_time = fb_obj.created_time
        m_likes = fb_obj.likes
        m_user_likes = fb_obj.user_likes
        m_type = fb_obj.type
        on error goto 0
        err.clear
    end sub
    
    
end class



Class List
	Private m_arraySize
	Private m_arrayData()
	Private m_Count
    
	Private Sub Class_Initialize
		m_arraySize=100
		Redim m_arrayData(m_arraySize-1)
		m_Count=0
	End Sub
	
	Public Property Get Count
		Count=m_Count
	End Property

	Public Sub Add(value)
    
		If m_Count>UBound(m_arrayData) Then
			Call Expand
		End If
        
		If IsObject(value) Then
			Set m_arrayData(m_Count)=value
		Else  
			m_arrayData(m_Count)=value
		End If
        
		m_Count=m_Count+1
	End Sub
	
	Public Default Property Get Item(index)
		
        If Not(IsNumeric(index)) Then
			Err.Raise 15000, "List", "invalid parameter given to Item: not numeric"
		End If
		If (index<0) Or (index>=m_Count) Then
			Err.Raise 15001, "List", "index out of range: (" & index & ") can't be less than zero or greater than count"
		End If
        
		If IsObject(m_arrayData(index)) Then
			Set Item=m_arrayData(index)
		Else  
			Item=m_arrayData(index)
		End If
        
	End Property
	
	Private Sub Expand
		ReDim Preserve m_arrayData(UBound(m_arrayData)+m_arraySize)
	End Sub

	Private Sub Class_Terminate
		Dim x
		For x=0 To UBound(m_arrayData)
			If IsObject(m_arrayData(x)) Then
				Set m_arrayData(x)=Nothing
			End If
		Next
		Erase m_arrayData
	End Sub
End Class
%>




<%

'' Copy and paste the template below for each new facebook dto class

'' A Facebook DTO
class fb_DTO

    '' **
    '' ** Declare Variables
    '' **
    
    '' Default Variables for properties 
    dim fbu '' Facebook utility object
    dim m_token ''
    dim m_json_str ''
    dim m_graph_url ''
    
    '' Default Internal Variables
    
    '' Object Type Specific Variables
    
    '' **   
    '' ** property exposure
    '' **

    
    '' Default Properties
    public property get graph_url()
        graph_url = m_graph_url
    end property
    
    public property get json_string()
        json_string = m_json_str
    end property
    
    public property get token()
        token = m_token
    end property
    public property let token(value)
        m_token = value
    end property
    
    
    '' Object Type property exposure
    
    
    '' initialization
    sub Class_Initialize
        set fbu = new fb_utility

    end sub
    
    '' termination
    sub Class_Terminate
        set fbu = nothing
        
    end sub

    '' **
    '' Implementation
    '' **
    
    '' Load object by id
    public sub Load(ID)
        dim obj
        
        m_graph_url = "https://graph.facebook.com/" & id & "?access_token=" & _
            m_token 

        m_json_str = fbu.get_page_contents( m_graph_url )
        
        
        
        set obj = JSON.parse( m_json_str )
        
        load_from_json_object obj
        
    end sub
    
    public sub load_from_json_object( fb_obj )
    
        on error resume next
        
        
        on error goto 0
        err.clear
    end sub
    
    
end class

%>
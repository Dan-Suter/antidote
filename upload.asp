
<!--#include virtual="/admin/freeASPUpload.asp"-->
<!--#include virtual="/admin/clsUpload.asp"-->
<!--#include virtual="/functions.asp"-->
<%
x=rw("uploading "&request.form("file"))
if request.form("file")<>"" then
	
end if
x=SaveFiles("../poeple/images/")


function SaveFilesCls()
set o = new clsUpload
'if o.Exists("cmdSubmit") then

'get client file name without path
sFileSplit = split(o.FileNameOf("file"), "\")
sFile = sFileSplit(Ubound(sFileSplit))

o.FileInputName = "file"
o.FileFullPath = Server.MapPath(".") & "\" & sFile
o.save

 if o.Error = "" then
	response.write "Success. File saved to  " & o.FileFullPath & ". Demo Input = " & o.ValueOf("Demo")
 else
	response.write "Failed due to the following error: " & o.Error
 end if

'end if
set o = nothing
end function

function SaveFiles(path)
    Dim Upload, fileName, fileSize, ks, i, fileKey
    Set Upload = New FreeASPUpload
    Upload.Save(path)

	' If something fails inside the script, but the exception is handled
	If Err.Number<>0 then
		x=rwb("error occured") 
		Exit function
	end if
    SaveFiles = ""
    ks = Upload.UploadedFiles.keys
    if (UBound(ks) <> -1) then
        SaveFiles = "<B>Files uploaded:</B> "
        for each fileKey in Upload.UploadedFiles.keys
            SaveFiles = SaveFiles & Upload.UploadedFiles(fileKey).FileName & " (" & Upload.UploadedFiles(fileKey).Length & "B) "
        next
    else
        SaveFiles = "The file name specified in the upload form does not correspond to a valid file in the system."
    end if
end function
%>







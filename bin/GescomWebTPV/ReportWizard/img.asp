<%
' ------------------------------------------------------------------------------------------
' img.asp
'
' To retrieve the content of the image.  There's no way to avoid creating a proxy page when
' pulling an image from a database.  However, this proxy page, designed by Tanwani Anyangwe (see
' comment in default.asp), accesses an existing recordset instead of executing a brand new query. 
' ------------------------------------------------------------------------------------------

%>
<% 
    Response.Expires = 0
    response.Buffer  = True
    response.Clear
   
    Response.ContentType = Session("ImageType")
    Response.BinaryWrite Session("ImageBytes")

    Session("ImageType") = ""
    Session("ImageBytes") = ""

    Response.End
%>
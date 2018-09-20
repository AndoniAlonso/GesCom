

<% 

Const ForReading = 1, ForWriting = 2, ForAppending = 8
dim fso, fs,filespec
set fso = server.createobject("Scripting.FileSystemObject") 
filespec = Server.MapPath("reports.txt")

if isobject(fso) then
	set fs = fso.opentextfile(filespec)
end if

'on error resume next

if request("a")="d" then
	dim ln,tempfile,s
	tempfile = Server.MapPath("tempfile.txt")
	Set f = fso.OpenTextFile(tempfile, ForAppending, True)		
		
	ln = CInt(request("ln"))
	while not fs.atendofline 
		s =  fs.readline
		If fs.Line <> ln Then
			f.WriteLine s
		End If		
	wend 
	
	fs.close
	set fs = nothing
	fso.DeleteFile filespec,True				
	
	f.Close		
	Set f = nothing
	
	fso.CopyFile tempfile, filespec,True
	fso.DeleteFile tempfile,True	
	set fso = nothing
	
	Response.Redirect ("default.asp")
else
	response.write("<b>Saved Reports</b><ul>")
	dim qs
	while not fs.atendofline 
		qs = Split(fs.readline,"|")
		response.write("<li><a href='default.asp?" & qs(1) & " '>" & qs(0) & "</a>&nbsp;")
		response.write("<a class=delete href='reports.asp?a=d&amp;ln=" & fs.Line &" '>x</a></li>")
	wend 
	response.write("</ul>")
	
	fs.close
	set fs = nothing
	set fso = nothing

end if
%> 

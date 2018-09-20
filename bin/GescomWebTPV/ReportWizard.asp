<!--#include file="dsn.asp"-->
<!--#include file="adovbs.inc" -->
<%
    'On Error Resume Next
    Response.Buffer = True
    
	'Get your table
	Dim strTable,strUrl,strFields,strCriteria, intPageSize, intCurrentPage
	Dim strReportName, blnSaveReport
	Dim sfield,stype,sdesc,sdesc2
	Dim sTotales, sFunciones, sDescripciones
	strTable = Request("t")
	strFields = Request("f")
	strCriteria = Request("c")
	strReportName = Request("rpt")
	blnSaveReport = False
	
	sfield = Request("sf")
	stype = Request("st")
	sdesc = Request("sd")
	sdesc2 = Request("sd2")
	
	sTotales = Split(sfield,"|")
	sFunciones = Split(stype,"|")
	sDescripciones = Split(sdesc,"|")
	
	intPageSize = Request("ps")
	intCurrentPage = Request("cp")
	
	strUrl = Request.ServerVariables("URL")
	
%>
<html>
	<head>
		<LINK href="./styles/GescomTPV.css" type="text/css" rel="stylesheet">
		<meta name="keywords" content="ASP,VBScript,Dynamic Reports">
		<meta name="author" content="Tanwani Anyangwe">
		<title>ASP Report Wizard</title>
		<SCRIPT LANGUAGE="JavaScript" SRC="TPVCommon.js"></SCRIPT>
		<script language=javascript>
		function OpenForm(table,fields,pagesize){
			var w = 500;
			var h = 500;
			var LeftPosition = (screen.width) ? (screen.width-w)/2 : 0;
			var TopPosition = (screen.height) ? (screen.height-h)/2 : 0;
			var settings = 'height='+h+',width='+w+',top='+TopPosition+',left='+LeftPosition+',scrollbars=0,resizable=1'
			var win = window.open("query_form.asp?t="+ table + "&f=" + fields + "&ps=" + pagesize,"query_form",settings)
		}
		</script>
		<style>
		A.colHeader
		{
		    FONT-SIZE: 12px;
		    COLOR: black;
		    FONT-FAMILY: Arial, Verdana, Arial;    
		    TEXT-DECORATION: none;
		}
		A.colHeader:visited
		{
		    FONT-SIZE: 12px;
		    COLOR: black;
		    FONT-FAMILY: Arial, Verdana, Arial;		    
		    TEXT-DECORATION: none;
		}
		A.colHeader:hover
		{
		    COLOR: black;
		    TEXT-DECORATION: underline;
		}
		A.delete{
			FONT-SIZE: 11px;
			FONT-WEIGHT:BOLD;
		    COLOR: red;
		    FONT-FAMILY:  Verdana, Arial;    
		    TEXT-DECORATION: none;
		}
		</style>
	</head>
	<body>		
		<%if strTable="" And Request.Form.Count < 1 then%>
		<h1>ASP Report Wizard</h1>
		<p>
			Welcome to Report Wizard.<br> 
			Please select a table below to generate your report			
		</p>
		<table>
		<tr><td>
		
		<%		
			'you can either list your tables here in the form
			Response.Write ("<b>Informes</b><ul>")
			Response.Write ("<li><a href='ReportWizard.asp?t=vTPVCierreCaja'>Cierre de Caja</a></li>")
			Response.Write ("<li><a href='ReportWizard.asp?t=vTPVReportVentas'>Listado de ventas</a></li>")
			Response.Write ("</ul>")
			
'			'or build your tables dynamically as follows
'			dim objCat,objTable
'			'Create a db catalog object
'			Set objCat = Server.CreateObject("ADOX.Catalog")
'			
'			'Point your catalog to you database
'			objCat.ActiveConnection = objConn
'   
'			'print you tables			
'			Response.Write ("<b>Tables</b><ul>")
'			For Each objTable In objCat.Tables		
'				'display only the user-created tables 		
'			    If UCase(objTable.Type) = "TABLE" Then
'			        Response.Write ("<li><a href=""" & strUrl & "?t=" & objTable.Name & """ >" & objTable.Name & "</a></li>")
'			    End If
'			Next
'			Response.Write ("</ul>")
'			
'			'clear objects from memory
'			Set objCat = Nothing
'			Set objTable = Nothing	
			%>
			</td>
			<td valign=top> 
			<!--#include file="reports.asp"-->
			</td></tr>
			</table>
			
			<%
		else
			if Request.Form("fields") <>"" then			
				strFields = Request.Form("fields")
				strCriteria =  Request.Form("criteria")	
				intPageSize = Request.Form("pagesize")		
				strReportName = Request.Form("report_name")
				
				blnSaveReport = CInt(Request.Form("save_report"))
				
				sfield = Request.Form("sfield")
				stype = Request.Form("stype")
				sdesc = Request.Form("sdesc")
				sdesc2 = Request.Form("sdesc2")
	
				if Len(Request.Form("table"))>0 then
					strTable =  Request.Form("table")
				end if	
			End if
			
			If strReportName="" Then
				strReportName = strTable & " Report"
			End If
					
					
			if Not IsNumeric(intPageSize) or intPageSize = "" then
				intPageSize = 10
			end if
			
			if Not IsNumeric(intCurrentPage) or intCurrentPage = "" then
				intCurrentPage = 1
			end if
			
			intPageSize = CInt(intPageSize)
			intCurrentPage = CInt(intCurrentPage)
			
			%>
  <table id="Container" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
  <tr>
	 <td vAlign="top" width="20%">&nbsp;
	 	<!--#include file="TPVMain.ASP"-->
   </td>
	<td vAlign="top">
			<h2><%=strReportName%></h2>
			<%	
			
			If strFields <> "" And strTable <> "" then	

				Dim strSQL, objRS, i
				strSQL = "SELECT "& strFields & " FROM [" & strTable & "]"
				If strCriteria <> "" Then
					strSQL = strSQL & " WHERE " & strCriteria
				End If
				
				'build sort order
				dim sort,lastsort,thissort
				sort = request("sort")
				lastsort = request("lastsort")

				if sort<>"" then
					if lastsort=sort then
						thissort = sort & " desc"
					elseif instr(lastsort,sort & " desc") then
						thissort = replace(lastsort,sort & " desc",sort)
					elseif instr(lastsort,sort) then
						thissort = replace(lastsort,sort,sort & " desc")
					elseif lastsort<>"" then
						thissort = lastsort & "," & sort
					else
						thissort = sort
					end if
					strSQL = strSQL & " ORDER BY " & thissort
					
					'Response.Write "<p><b><font color=blue>ORDER BY</font>:</b> " & thissort & "</p>"
					'Response.Write "<div><a href=""" & strUrl & "?t=" & strTable & "&amp;f="& strFields &"&amp;c="& strCriteria & """>Reset Order</a></div>"

				end if

				'Response.Write Replace(Replace(strSQL,"_perc","%"),"_qt","'")
				'Response.End
				
				Dim strQueryTable,strQuerySort,strQueryPage
				strQueryTable = "t=" & strTable & "&amp;f="& strFields &"&amp;c="& strCriteria & "&amp;sf="& sfield &"&amp;st="& stype &"&amp;sd="& sdesc &"&amp;sd2="& sdesc2 &"&amp;rpt="& strReportName
				strQueryPage = "ps="& intPageSize &"&amp;cp="& intCurrentPage				
				strQuerySort = "sort="& request("sort") &"&amp;lastsort="& request("lastsort") 
				
				If blnSaveReport Then
					SaveReport strReportName, strQueryTable & "&amp;" & strQuerySort & "&amp;cp="& i &"&amp;ps="& intPageSize
				End if	
				
				' Create recordset and set the page size
				Set objRS = Server.CreateObject("ADODB.Recordset")
				objRS.PageSize = intPageSize

				' You can change other settings as with any RS
				objRS.CursorLocation = adUseClient
				objRS.CacheSize = intPageSize

				' Open RS
				objRS.Open Replace(Replace(strSQL,"_perc","%"),"_qt","'"), objConn, 3,3'adOpenStatic, adLockReadOnly, adCmdText

				' Get the count of the pages using the given page size
				dim intPageCount
				intPageCount = objRS.PageCount

				' If the request page falls outside the acceptable range,
				' give them the closest match (1 or max)
				If intCurrentPage > intPageCount Then intCurrentPage = intPageCount
				If intCurrentPage < 1 Then intCurrentPage = 1

				' Check page count to prevent bombing when zero results are returned!
				If intPageCount = 0 Then
					Response.Write "<p>No hay ningún dato, varíe los criterios de búsqueda.</p>"
				Else
				
					' Move to the selected page
					objRS.AbsolutePage = intCurrentPage

					' Start output with a page x of n line
					Response.Write "<p>Página <b>" & intCurrentPage & "</b> de <b>" & intPageCount & "</b></p>"

					Response.Write "<table border=0 cellpadding=3 cellspacing=1 bgcolor=gray>"
				
					'add column headers
					Response.Write "<tr style=""background-color:#9999cc;"">"
					Response.Write "<th><a class=colHeader href=""" & strUrl &"?"& strQueryTable & """ title=""Reset Grid""><font face='ms outlook'>C</font></a></th>"
					For i=0 To objRS.Fields.Count - 1                
					    Response.Write "<th><a class=colHeader href=""" & strUrl & "?" & strQueryTable & "&amp;" & strQueryPage & "&amp;sort="& (i+1) &"&lastsort="& thissort & """>" & objRS.Fields(i).Name
						if instr(thissort,(i+1) & " desc") then
							Response.Write " -"
						elseif instr(thissort,(i+1) & "") then
							Response.Write " +"	
						end if
						Response.Write "</th>"	    
					Next
					Response.Write "</tr>"
				
					'add rows
					Dim intCounter,strColor
					DIm oleHeadSize
					oleHeadSize = 78
					
					Do While intCounter < intPageSize And  Not objRS.Eof
					    intCounter = intCounter + 1				    
					    'set the row color
					    If intCounter Mod 2 = 0 Then strColor = "white" Else strColor="gainsboro"
					    Response.Write "<tr style='background-color:"& strColor &";font-family:verdana;font-size:9pt;'>"
					    
					    'Add a row number
					    Response.Write "<td>" & ( intCounter + ( (intCurrentPage -1) * intPageSize) ) & "</td>"
					    
					    'loop thru all fields and display their values
					    For i=0 To objRS.Fields.Count - 1  
							Response.Write("<td>")
							Select Case objRS.Fields(i).Type
								Case 205:'Blob
									Call SetImgForDisp(objRS.Fields(i), "ole")
									Response.Write "<img src=""img.asp"" border=0>"
								Case Else : Response.Write(objRS.Fields(i).Value)
							End Select          
					        Response.Write("&nbsp;</td>")            
					    Next
					    				    
					    Response.Write "</tr>"
					    objRS.MoveNext
					Loop
					
					If stype<>"" and sfield<>"" Then
						dim rsTemp
						if sdesc="" then sdesc = sfield  & " " & sdesc2 
						'If strCriteria <> "" Then
						'	Set rsTemp = objConn.Execute("SELECT " & stype & "(" & sfield & ")" & " FROM " & strTable & " WHERE " & strCriteria)
						'Else
						'	Set rsTemp = objConn.Execute("SELECT " & stype & "(" & sfield & ")" & " FROM " & strTable)
						'End IF

						'Response.Write("<tr style=""background-color:#ffffe0;font-family:comic sans ms;font-size:11px;""><td colspan="& (objRS.Fields.Count + 1) &"><p><font face=arial><b>TOTAL</b></font><br>")
						'Response.Write( sdesc & ": " & Round(rsTemp(0),2) )
						'Response.Write("</p></td></tr>")
						dim sSQL, sSeparador
						sSQL = vbNullString 
						sSeparador = vbNullString 
						For i = 0 To UBound(sTotales) Step 1
							Dim sFuncionTotal,sCampoTotal
							sFuncionTotal = sFunciones(i) 
							sCampoTotal = sTotales(i) 
							sSQL = sSQL & sSeparador & sFuncionTotal & "(" & sCampoTotal & ")"
							sSeparador = ", "
						Next

						If strCriteria <> "" Then
							Set rsTemp = objConn.Execute("SELECT " & sSQL & " FROM " & strTable & " WHERE " & strCriteria)
						Else
							Set rsTemp = objConn.Execute("SELECT " & sSQL & " FROM " & strTable)
						End If
						
						Response.Write("<tr style=""background-color:#ffffe0;font-family:comic sans ms;font-size:11px;""><td colspan="& (objRS.Fields.Count + 1) &"><p><font face=arial><b>TOTALES</b></font><br>")
						For i = 0 To UBound(sTotales) Step 1
							Response.Write( sDescripciones(i) & ": " & Round(rsTemp(i),2) )
							Response.Write("<BR>")
						Next
						Response.Write("</p></td></tr>")
						
						rsTemp.Close
						Set rsTemp = Nothing
					End If
					
					Response.Write "</table>"

					'close recordset
					objRS.Close    
					Set objRS = Nothing
				
				    %><p><font face=verdana size=2><%
					' Show "previous" and "next" page links which pass the page to view
					' and any parameters needed to rebuild the query. 
					If intCurrentPage > 1 Then
						Response.Write "<a href=""" & strUrl & "?" & strQueryTable & "&amp;" & strQuerySort & "&amp;cp="& (intCurrentPage-intPageSize) &"&amp;ps="& intPageSize &""">[&lt;&lt; Primera]</a>&nbsp;"
					End If

					' You can also show page numbers:					
					Dim intStart, intStop
					
					intStop = intCurrentPage + 10 
					If intStop > intPageCount Then intStop = intPageCount
					
					intStart = intStop - 10
					If intStart < 1 Then intStart = 1
										
					
					For i = intStart To intStop										
						If i = intCurrentPage Then
							Response.Write(i & "&nbsp;") 
						Else
							Response.Write "<a href=""" & strUrl & "?" & strQueryTable & "&amp;" & strQuerySort & "&amp;cp="& i &"&amp;ps="& intPageSize &""">"& i &"</a>&nbsp;"
						End If
					Next 'I

					If intCurrentPage < intPageCount Then
						Response.Write "<a href=""" & strUrl & "?" & strQueryTable & "&amp;" & strQuerySort & "&amp;cp="& (intCurrentPage+intPageSize) &"&amp;ps="& intPageSize &""">[Ultima &gt;&gt;]</a>&nbsp;"
					End If
					
					'add exports
					dim strExportUrl
					strExportUrl = "export.asp?" & strQueryTable & "&amp;" & strQuerySort
					%>
					Descargar listado: 
					[&nbsp;
					<a href="<%=strExportUrl%>&amp;fm=1">Excel</a>&nbsp;|&nbsp;
					<a href="<%=strExportUrl%>&amp;fm=2">CSV</a>&nbsp;|&nbsp;
					<a href="<%=strExportUrl%>&amp;fm=3">XML</a>&nbsp;|&nbsp;
					<a href="<%=strExportUrl%>&amp;fm=4">Text</a>&nbsp;
					]	
					</font></p>			
	</td>
</tr>
</table>
					<%
				End If
				
			End if
			%>
		<%end if%>
	</body>
</html>
<%
	'close connection
    objConn.Close    
    Set objConn = Nothing
    
Function SetImgForDisp(inField, inContentType)
    Dim mOleHeadSize, mBytes
    Dim mFieldSize, mDumpaway
    mOleHeadSize = 78
    inContentType = LCase(inContentType)
    Select Case inContentType
        Case "gif", "jpeg", "bmp"
           inContentType = "image/" & inContentType
           mBytes = inField.value 

        Case "ole"
           inContentType = "image/bmp"  
           mFieldSize = inField.ActualSize
           mDumpaway = inField.GetChunk(mOleHeadSize)
           mBytes = inField.GetChunk(mFieldSize - mOleHeadSize)
    End Select
    Session("ImageBytes") = mBytes
    Session("ImageType") = inContentType
End Function

Function SaveReport(strName, strQuery)
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim fso, f, filespec
	filespec = Server.MapPath("reports.txt")
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.OpenTextFile(filespec, ForAppending, True)	
	f.WriteLine strName &  "|" & strQuery
	f.Close
	Set f = nothing
	set fso = nothing	
End Function



%>
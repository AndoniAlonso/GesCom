<!--#include file="dsn.asp"-->
<!--#include file="functions.asp"-->
<%
	Dim strTable,strFields,strCriteria,strSort,strSQL,strReportName
	strTable = Request("t")
	strReportName = Request("rpt")
	strFields = Request("f")
	strCriteria = Request("c")
	strSort = Request("sort")
	
	strSQL = "SELECT "& strFields & " FROM [" & strTable & "]"
	If strCriteria <> "" Then
		strSQL = strSQL & " WHERE " & strCriteria
	End If
	If strSort <> "" Then
		strSQL = strSQL & " ORDER BY " & strSort
	End If		
	
	Dim objRS, i,j		
	Set objRS = objConn.Execute(strSQL)		
	If objRS.EOF Then
		Response.Write "No records returned"
		Response.End 
	End if
	'objRS.Save "C:\dummy2.xml", 1
	Dim aryData, aryHeaders
	ReDim aryHeaders(objRS.Fields.Count)
	Dim intRecFirst, intRecLast
	Dim intFieldFirst, intFieldLast, intRecordCount 

	For i=0 To objRS.Fields.Count - 1                
	    aryHeaders(i) = objRS.Fields(i).Name			
	Next
	aryData = objRS.GetRows
	
	'empty recordset
	objRS.Close    
	Set objRS = Nothing
		
	'Find the Upper & Lower ends of the Array
	intRecFirst   = LBound(aryData, 2)
	intRecLast    = UBound(aryData, 2)
	intFieldFirst = LBound(aryData, 1)
	intFieldLast  = UBound(aryData, 1)
	intRecordCount = UBound(aryData, 2) + 1

    Dim fileName, format, contentType
    format = Cint(request("fm"))
    
    'set your variables and write data
    select case format
		case 1:PrintExcel'EXCEL
		case 2:PrintCSV 'CSV			
		case 3:PrintXML'XML			
		case else: PrintText 'text			
    end select 
                
	Response.End()
	
	Sub SetHeaders(fileName, contentType)
		'add your headers
		Response.ContentType = contentType 
		Response.AddHeader "Content-Disposition", "attachment; filename=" & fileName
	End Sub
	
	Sub PrintExcel
	
		SetHeaders strReportName & ".xls", "Application/vnd.excel"
		
		Response.Write "<html xmlns:x=""urn:schemas-microsoft-com:office:excel"">"
		Response.Write "<head>"
		Response.Write "<!--[if gte mso 9]><xml>"
		Response.Write "<x:ExcelWorkbook>"
		Response.Write "<x:ExcelWorksheets>"
		Response.Write "<x:ExcelWorksheet>"
		Response.Write "<x:Name>"& strTable &" Report</x:Name>"
		Response.Write "<x:WorksheetOptions>"
		Response.Write "<x:Print>"
		Response.Write "<x:ValidPrinterInfo/>"
		Response.Write "</x:Print>"
		Response.Write "</x:WorksheetOptions>"
		Response.Write "</x:ExcelWorksheet>"
		Response.Write "</x:ExcelWorksheets>"
		Response.Write "</x:ExcelWorkbook>"
		Response.Write "</xml>"
		Response.Write "<![endif]--> "
		Response.Write "</head>"
		Response.Write "<body>"
		
		PrintTable "<td>","</td>","<tr>","</tr>","<table>","</table>"
		
		Response.Write "</body>"
		Response.Write "</html>" 
	End Sub
	
	Sub PrintCSV
		SetHeaders strReportName & ".csv","text/csv" 
		PrintTable Chr(34),Chr(34) & Chr(44),"",Chr(10),"",""
	End Sub
	
	Sub PrintText
		SetHeaders strReportName & ".txt","text/txt" 
		PrintTable "",vbTab,"",vbCrlf,"",""
	End Sub
	
	Sub PrintXML
		SetHeaders strReportName & ".xml","text/xml"
		
		'Declare local variables.
		Dim objDom
		Dim objRoot
		Dim objField
		Dim objFieldValue
		Dim objcolName
		Dim objattTabOrder
		Dim objPI
		Dim x		
		Dim objRow

		'Instantiate the Microsoft XMLDOM.
		Set objDom = Server.CreateObject("Microsoft.XMLDOM")
		objDom.preserveWhiteSpace = True

		'Create your root element and append it to the XML document.
		 Set objRoot = objDom.createElement("root")
			objDom.appendChild objRoot

			'Iterate through each row in the Recordset
			For i = intRecFirst To intRecLast
					
				'Create a row-level node 
				Set objRow = objDom.CreateElement("record")		    
								    
				'loop thru all fields and display their values
				For j = intFieldFirst To intFieldLast  
						
					'*** Create an element, "field". ***
					Set objField = objDom.createElement("field")

					'*** Append the name attribute to the field node ***
					Set objcolName = objDom.createAttribute("name")
					objcolName.Text = aryHeaders(j)
					objField.SetAttributeNode(objColName)
					'***************************************************

					'*** Create a new node, "value". ***
					Set objFieldValue = objDom.createElement("value")

				'Set the value of the value node equal to the value of the
				'current field object				
				objFieldValue.Text = aryData(j,i)
				'************************************

				'*** Append the value node as a child of the field node. ***
				objField.appendChild objFieldValue
				'***********************************************************

				'*** Append the field node as a child of the row-level node. ***
				objRow.appendChild objField
				'***************************************************************
			Next 

			'*** Append the row-level node to the root node ***
			objRoot.appendChild objRow
			'**************************************************

		Next

		'*** Add the <?xml version="1.0" ?> tag ***
		Set objPI = objDom.createProcessingInstruction("xml", "version='1.0'")
		 
		'Append the processing instruction to the XML document.
		objDom.insertBefore objPI, objDom.childNodes(0)
		'************************************************

		'Write the XML contents as a string
		Response.Write(objDom.xml)

		'Clean up...
		Set objDom = Nothing
		Set objRoot = Nothing
		Set objField = Nothing
		Set objFieldValue = Nothing
		Set objcolName = Nothing
		Set objattTabOrder = Nothing
		Set objPI = Nothing
	End Sub
	
	Sub PrintTable(f1,f2,r1,r2,t1,t2)
		Response.Write t1	
					
		'add column headers
		Response.Write r1
		For i=0 To UBound(aryHeaders)                
		    Response.Write f1 & aryHeaders(i) & f2			
		Next
		Response.Write r2
				
		'add rows	
		For i = intRecFirst To intRecLast
		    Response.Write r1			    
					    
		    'loop thru all fields and display their values
		    For j = intFieldFirst To intFieldLast				      
		        Response.Write f1 & aryData(j, i) & f2            
		    Next
					    				    
		    Response.Write r2		    
		Next
		
		'close	
		Response.Write t2		
	End Sub
	
	
%>


<!--#include file="dsn.asp"-->
<!--#include file="functions.asp"-->
<html>
	<head>
		<meta name="keywords" content="ASP,VBScript,Dynamic Reports">
		<meta name="author" content="Tanwani Anyangwe">
		<title>Report Query Form</title>
		<script language=javascript>
		function ReturnVals(){
			var f = document.f1;
			setCriteria();
			if(f.fields.value.substring(0,1)==","){
				f.fields.value = f.fields.value.substring(1);
			}
			
			if(f.fields.value.length > 0 ){
				var fx = window.opener.document.forms[0];
				fx.fields.value = f.fields.value				
				fx.pagesize.value = f.pagesize.value
				fx.criteria.value = f.criteria.value
				fx.sfield.value = f.sfield.options[f.sfield.selectedIndex].value
				fx.stype.value = f.stype.options[f.stype.selectedIndex].value
				fx.sdesc.value = f.sdesc.value
				fx.sdesc2.value = f.stype.options[f.stype.selectedIndex].text
				fx.report_name.value = f.report_name.value
				
				if(f.save_report.checked){
					fx.save_report.value = 1;
				}else{
					fx.save_report.value = 0;
				}
				fx.submit();			
			}
			window.close();
		}
		
		function UpdateFields(name){
			var fields = document.f1.fields
			if(fields.value.indexOf(", [" + name + "]")>-1){
				fields.value = fields.value.replace(", [" + name + "]","")
			}else if(fields.value.indexOf("[" + name + "]")>-1){
				fields.value = fields.value.replace("[" + name + "]","")
			}else{
				fields.value = fields.value + ", [" + name + "]"
			}
		}
		
		function setCriteria(){
			var f = document.f1;
			if(f.search_field.selectedIndex>0 && 
				f.search_condition.selectedIndex>0 &&
				f.search_text.value!="" ){
				var s;				
				switch(f.search_condition.selectedIndex){
					case 1:s = " LIKE _qt_perc"+ f.search_text.value +"_perc_qt";break;
					case 2:s = " = _qt"+ f.search_text.value +"_qt";break;
					case 3:s = " < "+ f.search_text.value;break;
					case 4:s = " > "+ f.search_text.value;break;
					case 5:s = " LIKE _qt"+ f.search_text.value +"_perc_qt";break;
					case 6:s = " LIKE _qt_perc"+ f.search_text.value +"_qt";break;
				}
				f.criteria.value = f.search_field.options[f.search_field.selectedIndex].value + s;
			}
		}
		
		</script>
	</head>
	<body bgColor="#f5f5f5">
	<form name=f1>
	<table border=0  cellpadding=3 cellspacing=0>
	<tr><td  colspan=2 valign=top width=100%>
	<b>Select the fields you wish to display</b>
	
	<%
	Dim objRS,strTable,i,fname,options,pagesize,fields
	dim sdesc
	
	strTable = Request.QueryString("t")
	fields = Trim(Request("f"))
	pagesize = Request("ps")
	'Response.Write (pagesize)
	if pagesize="" or Not IsNumeric(pagesize) then
		pagesize = 10
	end if
	
	Set objRS = objConn.Execute("SELECT * FROM [" & strTable & "]")
	 
	Response.Write "<table border=0 valign=top><tr>"
	for i=0 to objRS.fields.count-1
		if i mod 4 = 0 then Response.Write "</tr><tr>"
		fname = objRS.fields(i).Name
	    Response.Write "<td><input type=checkbox onclick=""UpdateFields('" & fname & "')"" " & _
			" id=x" & i & " "& IIf(Instr(fields,"[" & fname & "]"),"checked","") &">&nbsp;<label for=x"& i & ">" & fname & "</label></td>"
		options = options & "<option value=""" & fname & """ >" & fname & "</option>"
	next
	Response.Write "</tr></table>"    
	
	%>
	</td>	
	</tr>
	<tr><td>
	<p><b>Records per page: </b> <input type=text name=pagesize value="<%= pagesize%>" size=5></p>
	</td>
	<td valign=top rowspan=2>	
	<input type=button value=" &#10; Generate &#10; Report &#10; " onclick="ReturnVals()" id=button1 name=button1>
	</td>	
	</tr>
	<tr><td>
	<p><b>Add any criteria for your report</b></p>
	<table><tr>
	<td>Field</td>
	<td>
	<select name="search_field">
	<option>&nbsp;</option>
	<%= options	%>
	</select>
	</td></tr>
	<tr><td>Condition</td>
	<td>
	<select name="search_condition">
		<option>&nbsp;</option>
		<option value="1">contains</option>
		<option value="2">is equal to</option>
		<option value="3">is less than</option>
		<option value="4">is greater than</option>
		<option value="5">starts with</option>
		<option value="6">ends with</option>
	</select>
	</td></tr>
	<tr><td>Value</td>
	<td>	
	<input type=text name=search_text value="">
	</td>	
	</tr></table>
	</td>
	
	</tr>
	
	<tr><td colspan=2>
	<p><b>Add a summary field to your report</b></p>
	<table><tr>
	<td>Field</td>
	<td>
	<select name="sfield">
	<option>&nbsp;</option>
	<%= options	%>
	</select>
	</td></tr>
	<tr><td>Summary type</td>
	<td>
	<select name="stype">
		<option>&nbsp;</option>
		<option value="SUM">Sum total</option>
		<option value="COUNT">Record Count</option>
		<option value="MIN">Minimum value</option>
		<option value="MAX">Maximum value</option>
		<option value="AVG">Average</option>
		<option value="STDEV">Standard deviation</option>
		<option value="VAR">Variance</option>
	</select>
	</td></tr>
	<tr><td>Summary Description</td>
	<td>	
	<input type=text name=sdesc value="<%=sdesc%>" size=40>
	</td>	
	</tr></table>
	</td>	
	</tr>
	
	<tr><td colspan=2>
	<p><b>
	<input type=checkbox name=save_report id=checkRpt>
	<label for=checkRpt>Save this report</label></b><br>	
	Report name: <input type=text name=report_name size=40>	
	</p>
	</tr>
	
	</table>
	<input type=hidden name=fields value="<%=fields%>">
	<input type=hidden name=criteria value="">
	</form>
	</body>
</html>
<%
	'close recordset
	objRS.Close    
	Set objRS = Nothing
	
	'close connection
    objConn.Close    
    Set objConn = Nothing
%>
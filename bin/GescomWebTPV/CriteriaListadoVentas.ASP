<html>
	<head>
		<SCRIPT LANGUAGE="JavaScript" SRC="TPVCommon.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript" SRC="CalendarPopup.js"></SCRIPT>
		<!-- This prints out the default stylehseets used by the DIV style calendar.
		     Only needed if you are using the DIV style popup -->
		<SCRIPT LANGUAGE="JavaScript">document.write(getCalendarStyles());</SCRIPT>
		<meta name="keywords" content="ASP,VBScript,Dynamic Reports">
		<meta name="author" content="Tanwani Anyangwe">
		<title>Report Query Form</title>
		<LINK href="./styles/GescomTPV.css" type="text/css" rel="stylesheet">
		<script language=javascript>
		function ReturnVals(){
			var f = document.f1;
			var report;
			setCriteria();
			
			report="ReportWizard.asp?t=vTPVReportVentas&f=%20[Fecha-Hora],%20[Numero],%20[Nombre%20vendedor],%20[Código%20artículo],%20[Descripcion],%20[Importe],%20[Nombre%20Centro]&c=" + f.criteria.value + "&sf=Importe&st=SUM&sd=Total&sd2=Sum%20total&rpt=Listado%20de%20ventas&sort=&lastsort=&cp=&ps=25";
			RedirectTo(report);			
		}
		
		function setCriteria(){
			var f = document.f1;
			// COMPROBAMOS QUE SE HAN RELLENADO LAS FECHAS
			//OJOOJO: validar que sea una fecha, etc.
			//OJOOJO: no habria que dejar seguir si no se ponen fechas correctas...
			if(f.dateDesde.value!="" &&
			   f.dateHasta.value!=""){
			   f.criteria.value = " [Fecha-Hora] BETWEEN '" + f.dateDesde.value + "' AND '" + f.dateHasta.value + "' ";
			   f.criteria.value = f.criteria.value + ' AND CentroGestionID = ' + f.CentroGestionID.value;
			   }
			else {
				f.criteria.value = "";
			}
		}
		
		</script>
	</head>
	<body>
	<form name=f1>
  	<table id="Container" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
  	<tr>
		<td vAlign="top" width="20%">&nbsp;
		 	<!--#include file="TPVMain.ASP"-->
	   </td>
		<td vAlign="top" >
			<table border=0  cellpadding=3 cellspacing=0>
			<tr><td  colspan=2 valign=top width=100%>
			<b>Listado de Ventas</b><BR>
			<b>Seleccionar los criterios del informe</b><BR><BR><BR>
			</td>	
			</tr>
			</table>
			<table border=0>
			<tr>
				<td>	
					<SCRIPT LANGUAGE="JavaScript" ID="jscalFechaDesde">
					var calFechaDesde = new CalendarPopup("testdiv1");
					//calFechaDesde.setCssPrefix("TEST");
					calFechaDesde.setMonthNames('Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'); 
					calFechaDesde.setDayHeaders('L','M','X','J','V','S','D'); 
					calFechaDesde.setWeekStartDay(1); 
					calFechaDesde.setTodayText("Hoy");
					var calFechaHasta = new CalendarPopup("testdiv1");
					//calFechaHasta.setCssPrefix("TEST");
					calFechaHasta.setMonthNames('Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'); 
					calFechaHasta.setDayHeaders('L','M','X','J','V','S','D'); 
					calFechaHasta.setWeekStartDay(1); 
					calFechaHasta.setTodayText("Hoy");
					</SCRIPT>
					Fecha desde:
				</td>	
				<td>	
					<INPUT TYPE="text" NAME="dateDesde" VALUE="<%Response.Write(Date())%>" SIZE=25>
					<A HREF="#" onClick="calFechaDesde.select(document.forms[0].dateDesde,'anchorDesde','dd/MM/yyyy'); return false;" TITLE="calFechaDesde.select(document.forms[0].dateDesde,'anchorDesde','MM/dd/yyyy'); return false;" NAME="anchorDesde" ID="anchorDesde">seleccionar</A>
				</td>	
			</tr>
			<tr>
				<td>Fecha hasta:
				</td>	
				<td>	
					<INPUT TYPE="text" NAME="dateHasta" VALUE="<%Response.Write(Date()+1)%>" SIZE=25>
					<A HREF="#" onClick="calFechaHasta.select(document.forms[0].dateHasta,'anchorHasta','dd/MM/yyyy'); return false;" TITLE="calFechaHasta.select(document.forms[0].dateHasta,'anchorHasta','MM/dd/yyyy'); return false;" NAME="anchorHasta" ID="anchorHasta">seleccionar</A>
				</td>	
			</tr>
			<tr>
				<td valign=top align=center>	
					<input type=button value=" &#10; Ejecutar &#10; informe &#10; " onclick="ReturnVals()" id=button1 name=button1>
				</td>	
			</tr>
			</table>
   	</td>
	</tr>
</table>
	<input type=hidden name=criteria value="">
	<input type=hidden name=CentroGestionID value="<%Response.Write(Request.Cookies("CentroGestionID"))%>">
	</form>
<DIV ID="testdiv1" STYLE="position:absolute;visibility:hidden;background-color:white;layer-background-color:white;"></DIV>
	</body>
</html>

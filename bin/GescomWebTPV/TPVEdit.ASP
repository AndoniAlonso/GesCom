<HTML>
  <HEAD>
		<title>Gescom TPV</title>
		<LINK href="./styles/GescomTPV.css" type="text/css" rel="stylesheet">

  </HEAD>
 <%	If Request.Cookies("TerminalID") = vbNullString Or Request.Cookies("CentroGestionID") = vbNullString Or Request.Cookies("AlmacenID") = vbNullString Then
 			Response.Redirect("./TPVIdentify.ASP")
 	End If	
 %>
 <%
   	Dim objAlbaranVenta
   	
   	Set objAlbaranVenta = Server.CreateObject("GescomObjects.AlbaranVenta")
   	' Si la varible de sesión que contiene el estado del ticket no tiene ningún valor, suponemos que se trata de crear un ticket nuevo
   	If VarType(Session("AlbaranVentaState")) <> vbArray + vbByte Then
			' Creamos el objeto parametro y lo cargamos
   		Dim objParametro 
   		Set objParametro = Server.CreateObject("GescomObjects.Parametro")
   		objParametro.Load

		objAlbaranVenta.BeginEdit
		objAlbaranVenta.Numero = objParametro.ObjEmpresaActual.IncrementaAlbaranVentas
   		objAlbaranVenta.TemporadaID = objParametro.TemporadaActualID
        objAlbaranVenta.EmpresaID = objParametro.EmpresaActualID
        objAlbaranVenta.TerminalID = Request.Cookies("TerminalID")
        objAlbaranVenta.CentroGestionID = Request.Cookies("CentroGestionID")
        objAlbaranVenta.AlmacenID = Request.Cookies("AlmacenID")
        objAlbaranVenta.AsignarClientePredeterminado
        Session("AlbaranVentaState")= objAlbaranVenta.GetSuperState
	Else
		objAlbaranVenta.SetSuperState Session("AlbaranVentaState")
   	End If
 %>
  <BODY onload="document.Form1.TPVEdit_txtBarCode.focus();document.onkeydown=DetectSpecialKeys;" 
 <% 	'
  	If objAlbaranVenta.EsDevolucion Then
  		Response.Write ("style='background-color: #CC6699;'")
  	End If
 %>			
  >
  <table id="Container" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
  <tr>
	 <td vAlign="top" width="20%">&nbsp;
	 	<!--#include file="TPVMain.ASP"-->
   </td>
	<td vAlign="top">
		<table cellSpacing="0" cellPadding="0" width="600" border="0">
			<tr>
				<td class="navigation">Punto de venta</td>
			</tr>
		</table>

      <table id="tblListe" width="600" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td><br></td>
			</tr>
			<tr>
				<td class="titreinter">Capturar código de barras</td>
			</tr>
			<tr>
				<td><br></td>
			</tr>
			<tr>
				<td>
					<form name="Form1" method="post" action="" id="Form1">
					<table id="tblListe" width="600" border="0" cellpadding="0" cellspacing="0">
						<tr class="TitreSection">
							<td>Código de barras<BR>Tecla F5</td>
							<td><input name="txtBarCode" type="text" id="TPVEdit_txtBarCode" class="form_obj" onkeypress="BarCodeKeyPress();" /></td>
							<td>Número</td>
							<td><% Response.Write (objAlbaranVenta.Numero)
								 %>
							</td>
						</tr>
					</table>
					</form>
				</td>
			</tr>
			<tr>
				<td>

					<table cellspacing="0" cellpadding="4" rules="cols" bordercolor="#DEDFDE" border="1" id="dgEvenements" style="color:Black;background-color:White;border-color:#DEDFDE;border-width:1px;border-style:None;border-collapse:collapse;">
						<tr style="color:white;background-color:#6B696B;font-weight:bold;">
							<td style="color:white">Artículo - color</td><td style="color:white">Talla</td><td style="color:white">Precio</td><td style="color:white">Descuento</td><td style="color:white">Cantidad</td><td style="color:white">Importe<BR>Tecla - F2</td><td style="color:white">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
						</tr>
						<% Dim i 
						   Dim objAlbaranVentaItem 
						   Dim bolImpares
						   bolImpares = True
						   
							For i = 1 To objAlbaranVenta.AlbaranVentaItems.Count
								Set objAlbaranVentaItem = objAlbaranVenta.AlbaranVentaItems(i)
						
								If bolImpares Then %>
						<tr nowrap="nowrap" onmouseover="this.style.backgroundColor='powderblue';" onmouseout="this.style.backgroundColor='#F7F7DE';" style="background-color:#F7F7DE;">
						<%		Else  				 %>
						<tr nowrap="nowrap" onmouseover="this.style.backgroundColor='powderblue';" onmouseout="this.style.backgroundColor='White';" style="background-color:White">
						<%		End If  				 
								bolImpares = not bolImpares
						%>
							<td>
								<% If objAlbaranVentaItem.ArticuloColorID <> 0 Then
										Response.Write objAlbaranVentaItem.ArticuloColor
									Else
									   Response.Write objAlbaranVentaItem.Descripcion
									End If
								%>
							</td>
							<td align=center>
								<%
					            Select Case True
					            Case objAlbaranVentaItem.CantidadT36 <> 0
					                Response.Write "36"
					            Case objAlbaranVentaItem.CantidadT38 <> 0
					                Response.Write "38"
					            Case objAlbaranVentaItem.CantidadT40 <> 0
					                Response.Write "40"
					            Case objAlbaranVentaItem.CantidadT42 <> 0
					                Response.Write "42"
					            Case objAlbaranVentaItem.CantidadT44 <> 0
					                Response.Write "44"
					            Case objAlbaranVentaItem.CantidadT46 <> 0
					                Response.Write "46"
					            Case objAlbaranVentaItem.CantidadT48 <> 0
					                Response.Write "48"
					            Case objAlbaranVentaItem.CantidadT50 <> 0
					                Response.Write "50"
					            Case objAlbaranVentaItem.CantidadT52 <> 0
					                Response.Write "52"
					            Case objAlbaranVentaItem.CantidadT54 <> 0
					                Response.Write "54"
					            Case objAlbaranVentaItem.CantidadT56 <> 0
					                Response.Write "56"
					            Case Else
					            ''ojoojo: devolver error
					            End Select
					         %>
							</td>
							<td align=right>
								<%
									Response.Write FormatNumber(objAlbaranVentaItem.PrecioVenta,2)
					         %>
							</td>
							<td align=right>
								<table border=0>
									<tr>
										<td rowspan=2 >
								<%
								
									Response.Write "<input name='txtDescuento" & i & "' type='text' align=center maxlength=3 size=4 id='TPVEdit_txtDescuento' onkeypress='DescuentoKeyPress();' "  
									Response.Write "value ='" & objAlbaranVentaItem.Descuento & "'>"
					         %>
										</td>
										<td onclick='UpDownDescuento(1,document.all.txtDescuento<%Response.Write(i)%>);' title='Incrementar el descuento - [Numpad +]'><IMG src='images/ico_mas.gif'></td>
									</tr>
									<tr>
										<td onclick='UpDownDescuento(-1,document.all.txtDescuento<%Response.Write(i)%>);' title='Decrementar el descuento - [Numpad -]'><IMG src='images/ico_menos.gif'></td>
									</tr>
								</table>
							</td>
							<td align=center>
								<%
									Response.Write objAlbaranVentaItem.Cantidad
					         %>
							</td>
							<td align=right>
								<%
									Response.Write "<input name='txtImporteBruto" & i & "' type='text' align=center maxlength=7 size=7 id='TPVEdit_txtImporteBruto' onkeypress='ImporteBrutoKeyPress();' "  
									Response.Write "value ='" & FormatNumber(objAlbaranVentaItem.BRUTO,2) & "'>"
					         %>
							</td>
							<td>
								<input type="image" name="ModificarLinea" id="ModificarLinea" title="Modificar este artículo - [Numpad /]" src="./images/edit.gif" alt="Editar" onclick="ModificarLinea(<%Response.Write(i)%>,document.all.txtDescuento<%Response.Write(i)%>.value,document.all.txtImporteBruto<%Response.Write(i)%>.value);" />
								<input type="image" name="EliminarLinea" id="EliminarLinea" title="Eliminar este artículo - Tecla F12" src="./images/delete.gif" alt="Eliminar" onclick="EliminarLinea(<%Response.Write(i)%>);"/>
							</td>
						</tr>
						<% Next
						%>
						<%		If bolImpares Then %>
						<tr nowrap="nowrap" onmouseover="this.style.backgroundColor='powderblue';" onmouseout="this.style.backgroundColor='#F7F7DE';" style="background-color:#F7F7DE;">
						<%		Else  				 %>
						<tr nowrap="nowrap" onmouseover="this.style.backgroundColor='powderblue';" onmouseout="this.style.backgroundColor='White';" style="background-color:White">
						<%		End If  				 %>
							<td>TOTAL TICKET
							</td>
							<td align=center></td>
							<td align=center></td>
							<td align=center></td>
							<td align=center></td>
							<td align=center>
								<%
									Response.Write FormatNumber(objAlbaranVenta.TotalBruto,2)
								%>
							</td>
							<td align=center>
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td>&nbsp;</td>
			</tr>
			<tr>
				<td>
					<table id="tblBotones" width="600" border="0" cellpadding="0" cellspacing="0">
						<tr class="TitreSection">
							<td><a href="javascript:FinTicket();" id="FinTicket" name="FinTicket"><DIV align=center><IMG src='images/cashier.png' width=48 height=48 border=0	><BR>Fin Ticket<BR>Tecla - F6</a></DIV></td>
							<td><a href="javascript:CancelarTicket();" id="CancelarTicket" name="CancelarTicket"><DIV align=center><IMG src='images/shoppingbasket_delete.png' width=48 height=48 border=0	><BR>Cancelar Ticket<BR>Tecla - F7</a></DIV></td>
							<td><a href="javascript:ImprimirUltimoTicket();"><DIV align=center><IMG src='images/printer_preferences.png' width=48 height=48 border=0	><BR>Imprimir último ticket<BR>Tecla - F8</a></DIV></td>
							<td><a href="javascript:Devolucion();"><DIV align=center><IMG src='images/replace2.png' width=48 height=48 border=0	><BR>Devolución<BR>Tecla - F9</a></DIV></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</td>
</tr>
</table>
		
<SCRIPT LANGUAGE="JavaScript" SRC="TPVCommon.js"></SCRIPT>
<SCRIPT>
<!--
// Enlazar con el alta de líneas
function BarCodeKeyPress ()
	{
	txtBarcode = new String(document.Form1.TPVEdit_txtBarCode.value);
	if (window.event.keyCode==13) { //Se ha pulsado enter??
		// truncamos el string si tiene más de 9 caracteres (por si pegan el código con CTRL + V)
		if (txtBarcode.length>9) 
			txtBarcode = txtBarcode.substring(0,8); 
		if (txtBarcode.length != 0) {
			window.event.keyCode = 0;
			RedirectTo('TPVEditDoActions.ASP?Action=AlbaranItemCodigoBarras&txtBarCode=' + txtBarcode);
		}
	}
	// sólo dejamos que se tecleen caracteres numéricos
	if (window.event.keyCode < 48 || window.event.keyCode > 57)	{
		window.event.keyCode=0;
	}
	// no dejar que tenga más de 9 caracteres
	if (txtBarcode.length > 9) {
		window.event.keyCode=0;
	}
   }

function DescuentoKeyPress()
	{
	// sólo dejamos que se tecleen caracteres numéricos
	if (window.event.keyCode < 48 || window.event.keyCode > 57)	{
		window.event.keyCode=0;
	}
	}

function ImporteBrutoKeyPress()
	{
	// sólo dejamos que se tecleen caracteres numéricos y la coma decimal
	// el punto lo transformamos en coma.
	if (window.event.keyCode == 46) {
		window.event.keyCode = 44;
	}
	if ((window.event.keyCode < 48 || window.event.keyCode > 57) && window.event.keyCode != 44)	{
		window.event.keyCode=0;
	}
	}

function EliminarLinea(numLinea)
	{
	if (window.confirm('Se va a eliminar la venta realizada del artículo..\nPara aceptar esta acción, haga clic en Aceptar, si no haga clic en Cancelar')) {
		RedirectTo('TPVEditDoActions.ASP?Action=EliminarLinea&Key=' + numLinea);
		}
	}

function ModificarLinea(numLinea,Descuento,ImporteBruto)
	{
	RedirectTo('TPVEditDoActions.ASP?Action=ModificarLinea&Key=' + numLinea + '&Descuento=' + Descuento + '&ImporteBruto=' + ImporteBruto);
	}

function FinTicket()
	{
	document.all.FinTicket.disabled=true;
	RedirectTo('TPVGeneralEdit.ASP');
	}
	
function CancelarTicket()
	{
	if (window.confirm('Se va a eliminar el ticket completamente..\nPara aceptar esta acción, haga clic en Aceptar, si no haga clic en Cancelar')) {
		RedirectTo('TPVEditDoActions.ASP?Action=CancelEdit');
		}
	}
	
function ImprimirUltimoTicket()
	{
	RedirectTo('TPVEditDoActions.ASP?Action=PrintLastTicket');
	}
	
function Devolucion()
	{
	RedirectTo('TPVEditDoActions.ASP?Action=AlternarDevolucion');
	}
	
function UpDownDescuento(SumaResta, InputDescuento) 
    {   var Lote = 10;
        var NuevoDescuento; 

        //NuevoDescuento = parseInt(document.all.txtDescuento.value) + Lote * SumaResta; 
        NuevoDescuento = parseInt(InputDescuento.value) + Lote * SumaResta; 
        if (NuevoDescuento >= 0 && NuevoDescuento <= 100) { 
                InputDescuento.value =  NuevoDescuento; 
        } 
        InputDescuento.focus();
     } 

function OpenCriteriaForm(CriteriaReportName)
	{	  var w = 500;
		  var h = 500;
		  var LeftPosition = (screen.width) ? (screen.width-w)/2 : 0;
		  var TopPosition = (screen.height) ? (screen.height-h)/2 : 0;
		  var settings = 'height='+h+',width='+w+',top='+TopPosition+',left='+LeftPosition+',scrollbars=0,resizable=1'
		  var win = window.open(CriteriaReportName,'CriteriaReportName',settings)
	}

function DetectSpecialKeys(e){
	//var evtobj=window.event? event : e;
	//if (evtobj.altKey || evtobj.ctrlKey || evtobj.shiftKey)
	//alert("you pressed one of the 'Alt', 'Ctrl', or 'Shift' keys");

	// F6 - Fin ticket
	if (window.event.keyCode == 117) { 
		window.event.keyCode = 0;
		FinTicket();
	}

	// F7 - Cancelar ticket
	if (window.event.keyCode == 118) { 
		window.event.keyCode = 0;
		CancelarTicket();
	}
	
	// F8 - Imprimir último ticket
	if (window.event.keyCode == 119) { 
		window.event.keyCode = 0;
		ImprimirUltimoTicket();
	}

	// F9 - Devolución
	if (window.event.keyCode == 120) { 
		window.event.keyCode = 0;
		Devolucion();
	}
	
	// F2 - Importe
	if (window.event.keyCode == 113) { 
		window.event.keyCode = 0;
		document.all.txtImporteBruto<%Response.Write(i-1)%>.focus();
		document.all.txtImporteBruto<%Response.Write(i-1)%>.select();
	}
	
	// Numpad +		
	if (window.event.keyCode == 107) { 
		window.event.keyCode = 0;
		UpDownDescuento(1,document.all.txtDescuento<%Response.Write(i-1)%>);
	}

	// Numpad -		
	if (window.event.keyCode == 109) { 
		window.event.keyCode = 0;
		UpDownDescuento(-1,document.all.txtDescuento<%Response.Write(i-1)%>);
	}

	// Numpad /		
	if (window.event.keyCode == 111) { 
		window.event.keyCode = 0;
		ModificarLinea(<%Response.Write(i-1)%>,document.all.txtDescuento<%Response.Write(i-1)%>.value,document.all.txtImporteBruto<%Response.Write(i-1)%>.value);
	}

	// F12			
	if (window.event.keyCode == 123) { 
		window.event.keyCode = 0;
		EliminarLinea(<%Response.Write(i-1)%>);
	}

	// CTRL + ALT + I 	permite cambiar de centro de conexión
	//if (window.event.keyCode == 123) { 
	//	window.event.keyCode = 0;
	//	EliminarLinea(<%Response.Write(i-1)%>);
	//}
}

//-->
</SCRIPT>

</BODY>
</HTML>

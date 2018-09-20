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
		objAlbaranVenta.SetSuperState Session("AlbaranVentaState")
 %>
  <BODY onload="document.all.cboRepresentantes.focus();document.onkeydown=DetectSpecialKeys;">
  <table id="Container" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
  <tr>
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
				<td class="titreinter">Datos del ticket de venta</td>
			</tr>
			<tr>
				<td><br></td>
			</tr>
			<tr>
				<td>
					<table id="tblListe" width="600" border="0" cellpadding="0" cellspacing="0">
						<tr class="TitreSection">
							<td>Cliente</td>
							<td>
								<%
									LoadCombo objAlbaranVenta.Clientes, objAlbaranVenta.Cliente, "cboClientes"
								%>
							</td>
						</tr>
						<tr class="TitreSection">
							<td>Vendedor</td>
							<td>
								<%
									LoadCombo objAlbaranVenta.Representantes, objAlbaranVenta.Representante, "cboRepresentantes"
								%>
							</td>
						</tr>
						<tr class="TitreSection">
							<td>Medio de pago</td>
							<td>
								<%
									LoadCombo objAlbaranVenta.MediosPago, objAlbaranVenta.MedioPago, "cboMediosPago"
								%>
							</td>
						</tr>
						<tr class="TitreSection">
							<td><BR></td>
							<td><BR></td>
						</tr>
						<tr class="TitreSection">
							<td>TOTAL TICKET</td>
							<td>
								<%
									Response.Write FormatNumber(objAlbaranVenta.TotalBruto,2) & " €"
								%>
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td>

				</td>
			</tr>
			<tr>
				<td>&nbsp;</td>
			</tr>
			<tr>
				<td>
					<table id="tblBotones" width="600" border="0" cellpadding="0" cellspacing="0">
						<tr class="TitreSection">
							<td><a href="javascript:AceptarCambios();" id="AceptarCambios" name="AceptarCambios"><DIV align=center><IMG src='images/check2.png' width=48 height=48><BR>Fin Ticket<BR>Tecla - F6</a></DIV></td>
							<td><a href="javascript:CancelarCambios();"><DIV align=center><IMG src='images/delete2.png' width=48 height=48><BR>Cancelar Ticket<BR>Tecla - F7</a></DIV></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</td>
</tr>
</table>
<INPUT type="hidden" id=objAlbaranVenta_IsNew name="objAlbaranVenta_IsNew" value="<%Response.Write(LCase(CStr(objAlbaranVenta.IsNew)))%>">	
<INPUT type="hidden" id=Representante_IsChanged name="Representante_IsChanged" value="false">	
<INPUT type="hidden" id=MedioPago_IsChanged name="MedioPago_IsChanged" value="false">	
<SCRIPT>
<!--
// Redireccionar a otras páginas (helper)
function RedirectTo(NewDirection)
    {
    window.location.href = NewDirection;
    }
    
function AceptarCambios ()
	{
		if (ValidarDatos())
		{
			var strCliente = document.all.cboClientes.value;
			var strRepresentante = document.all.cboRepresentantes.value;
			var strMedioPago = document.all.cboMediosPago.value;
			var dtFecha = new Date();
		
			var strDireccion = "TPVEditDoActions.ASP?Action=ApplyEdit" +
			                   "&Cliente=" + strCliente +
			                   "&Representante=" + strRepresentante +
			                   "&MediosPago=" + strMedioPago +
			                   "&Fecha" + dtFecha;
				                   
			document.body.innerText ="";    // Con esto evito que si se pulsa "doble click" intente emitir dos veces el ticket.
			RedirectTo(strDireccion);
		}
		else	
		{
			alert("Para finalizar el Ticket es necesario indicar el medio de pago y el vendedor.");
		}
 	}
 	
 function ValidarDatos ()
	{
		var Resultado = new Boolean();
		
		if (document.all.objAlbaranVenta_IsNew == 'false')
		{
			Resultado = true;
		}
		else
		{
			if (document.all.Representante_IsChanged.value=='true' && document.all.MedioPago_IsChanged.value=='true')
			{
				Resultado = true;
			}
			else
			{
				Resultado = false;
				document.all.cboRepresentantes.focus();
			}
		}
		
		return Resultado;
	}

function CancelarCambios(){
	RedirectTo('TPVEdit.ASP');
	}
	
function DetectSpecialKeys(e){
	//var evtobj=window.event? event : e;
	//if (evtobj.altKey || evtobj.ctrlKey || evtobj.shiftKey)
	//alert("you pressed one of the 'Alt', 'Ctrl', or 'Shift' keys");

	// F6 - Fin ticket
	if (window.event.keyCode == 117) { 
		window.event.keyCode = 0;
		AceptarCambios();
	}

	// F7 - Cancelar ticket
	if (window.event.keyCode == 118) { 
		window.event.keyCode = 0;
		CancelarCambios();
	}
	
}

function Change_cboClientes ()
	{
		//
	}

function Change_cboRepresentantes ()
	{
		document.all.Representante_IsChanged.value = true;
		document.all.cboMediosPago.focus();
	}

function Change_cboMediosPago ()
	{
		document.all.MedioPago_IsChanged.value = true;
		document.all.AceptarCambios.focus();
	}
//-->
</SCRIPT>
<%
	Public Sub LoadCombo(List, SelectedValue, ComboName)
		Dim vntItem 
	  
	  	Response.Write ("<SELECT NAME='" & ComboName & "' class='form_obj' onChange='Change_" & ComboName & "();'>")
							  	
	   For Each vntItem In List
	   	If SelectedValue = vntItem Then
	  	   	Response.Write ("<OPTION VALUE='" & vntItem & "' SELECTED>" & vntItem) 
	   	Else
		   	Response.Write ("<OPTION VALUE='" & vntItem & "'>" & vntItem) 
	   	End If
		Next
	
	  	Response.Write ("</SELECT>")

	End Sub
%>

</BODY>
</HTML>

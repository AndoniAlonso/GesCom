<HTML>
  <HEAD>
		<title>Ejecutar las acciones sobre un albaran de venta (un ticket)</title>
		<LINK href="./styles/GescomTPV.css" type="text/css" rel="stylesheet">

<OBJECT id="objPrintTicket" CLASSID="CLSID:FF2B1563-BF62-4666-B50A-509A2E9CA5CB">
</OBJECT>

<script language="JavaScript">
<!--
function ImpresionTicket ()
{
	if (window.print) {
	  text = document.mxForm.ImpresionTicket.value;
	  winId = window.open('','newwin');
	  with (winId.document) {
	    write('<body onLoad="window.focus();window.print();close();"><pre>'+text+'</pre><\/body>');
	    close();
	  }
	}
}

function fnPrintTicket ()
{
	var objPrintTicket;
	
   text = document.mxForm.ImpresionTicket.value;
	//window.Parent.document.objPrintTicket.PrintTicket(text);
	objPrintTicket = new ActiveXObject("WebBrowserPrint.PrintTicket"); 
	objPrintTicket.PrintTicket(text);
}
//-->
</script>
  </HEAD>
  <BODY onload="fnPrintTicket();window.location.href='TPVEdit.ASP';">
  	<FORM name="mxForm">
  		<input name="txtBarCode" type="hidden" id="ImpresionTicket" value='<%Response.Write(Session("ImpresionTicket"))%>'/>
  	</FORM>
  </BODY>
</HTML>

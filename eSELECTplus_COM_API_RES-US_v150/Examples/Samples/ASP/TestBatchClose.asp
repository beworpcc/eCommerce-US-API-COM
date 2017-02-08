<% @Language = "VBScript" %>
<% Response.buffer = true %>
<%
	store_id = "monusqa002"
	api_token = "qatoken"

	ecr_number="64000003"

	Set out = server.CreateObject("Moneris.USRequest")
	out.initRequest store_id, api_token, "https://esplusqa.moneris.com/gateway_us/servlet/MpgRequest"

	Set batchClose = server.CreateObject("Moneris.USBatchClose")

	out.setRequest batchClose.formatRequest(ecr_number)
	out.sendRequest
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" Content="text-html; charset=Windows-1252">
</HEAD>
<BODY bgcolor=white>

<%
	Response.Write "Receipt ID:  " & out.getReceiptID & "<br>"
	Response.Write "Response Code:  " & out.getResponseCode & "<br>"
	Response.Write "Transaction Type:  " & out.getTransType & "<br>"
	Response.Write "Message:  " & out.getMessage & "<br>"
	Response.Write "Amount:  " & out.getTransAmount & "<br>"
	Response.Write "Bank Totals:  " & out.getBankTotals & "<br>"
	Response.Write "Card Type:  " & out.getCardType & "<br>"
	Response.Write "Reference Number:  " & out.getReferenceNum & "<br>"
	Response.Write "Transaction ID:  " & out.getTransID & "<br>"
	Response.Write "ISO:  " & out.getISO & "<br>"
	Response.Write "Auth Code:  " & out.getAuthCode & "<br>"
	Response.Write "Transaction Time:  " & out.getTransTime & "<br>"
	Response.Write "Transaction Date:  " & out.getTransDate & "<br>"
	Response.Write "Complete:  " & out.getCompleteStatus & "<br>"
	Response.Write "Timeout:  " & out.getTimedoutStatus & "<br>"
	Response.Write "Ticket:  " & out.getTicket & "<br>"
	Response.Write "dumpXML:  " & out.dumpXMLResponse & "<br>"
%>

</BODY>
</HTML>

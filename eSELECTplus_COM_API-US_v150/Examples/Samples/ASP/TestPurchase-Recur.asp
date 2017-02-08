<% @Language = "VBScript" %>
<% Response.buffer = true %>
<%
	store_id = "monusqa002"
	api_token = "qatoken"

	order_id = "recur-"&Day(Date)&Month(Date)&Year(Date)&"-"&Hour(Now)&Minute(Now)&Second(Now)
	amount="5.00"
	pan="4242424242424242"
	exp_date = "1111"
	cust_id = "cust_id"
	crypt_type="7"
	dynamic_descriptor = "12345"

	'==================='
	'Recurring setup'
	recur_unit = "month"
	start_now = "true"
	start_date = "2011/11/01"
	num_recurs = "4"
	period = "1"
	recur_amount = "15.00"

	'======================================'
	'optional Level 2 details'
	commcard_invoice = "Invoice 56789"
	commcard_tax_amount = "0.08"

	Set out = server.CreateObject("Moneris.USRequest")
	out.initRequest store_id, api_token, "https://esplusqa.moneris.com/gateway_us/servlet/MpgRequest"

	Set request1 = server.CreateObject("Moneris.USPurchase")
		request1.setCustId cust_id
		request1.setDynamicDescriptor dynamic_descriptor
		request1.setRecur recur_unit,start_now,start_date,num_recurs,period,recur_amount
		request1.setCommcardInvoice( commcard_invoice )
		request1.setCommcardTaxAmount( commcard_tax_amount)

	out.setRequest request1.formatRequest(order_id, amount, pan, exp_date, crypt_type)
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
	Response.Write "RecurSuccess:  " & out.getRecurSuccess & "<br>"
%>

</BODY>
</HTML>

<% @Language = "VBScript" %>
<% Response.buffer = true %>
<%
	store_id = "monusqa002"
	api_token = "qatoken"

	order_id = "PinDebPurch-"&Day(Date)&Month(Date)&Year(Date)&"-"&Hour(Now)&Minute(Now)&Second(Now)
	amount="35.00"
	pan="4496270000164824"
	cust_id = "cust_id"

	presentation_type="W"
	intended_use="1"
	p_account_number="1234567890123456789012345"

	'==============
	'Recur Info
	recur_unit = "month"
	start_now = "true"
	start_date = "2011/11/01"
	num_recurs = "4"
	period = "1"
	recur_amount = "15.00"

	Set out = server.CreateObject("Moneris.USRequest")
	out.initRequest store_id, api_token, "https://esplusqa.moneris.com/gateway_us/servlet/MpgRequest"

	Set pinlessdebitpurchase = server.CreateObject("Moneris.USPinlessDebitPurchase")
		pinlessdebitpurchase.setCustId cust_id
		pinlessdebitpurchase.setRecur recur_unit,start_now,start_date,num_recurs,period,recur_amount

	out.setRequest pinlessdebitpurchase.formatRequest(order_id, amount, pan, presentation_type, intended_use, p_account_number)
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

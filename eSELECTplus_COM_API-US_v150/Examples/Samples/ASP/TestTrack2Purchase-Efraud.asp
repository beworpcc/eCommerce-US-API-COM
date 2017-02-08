<% @Language = "VBScript" %>
<% Response.buffer = true %>
<%
	store_id = "monusqa002"
	api_token = "qatoken"

	order_id = "t2purch-"&Day(Date)&Month(Date)&Year(Date)&"-"&Hour(Now)&Minute(Now)&Second(Now)
	amount="10.42"
	cust_id = "cust_id"
	track2=";4924190000004030=09121214797536211133?"
	pos_code="00"
	pan=""
	exp_date = ""
	crypt_type="7"

	'===================='
	'Optional Level 2/3 variables'
	commcard_invoice ="Invoice 123456"
	commcard_tax_amount="0.07"

	'==================
	'EFraud - AVS setup
	avs_street_number = "6600"
	avs_street_name = "New York Street"
	avs_zipcode = "90210"

	cvd_indicator = "1"
	cvd_value = "333"

	Set out = server.CreateObject("Moneris.USRequest")
	out.initRequest store_id, api_token, "https://esplusqa.moneris.com/gateway_us/servlet/MpgRequest"

	Set track2purchaserequest = server.CreateObject("Moneris.USTrack2Purchase")
		track2purchaserequest.setCustId cust_id
		track2purchaserequest.setAvsInfo avs_street_number, avs_street_name, avs_zipcode
		track2purchaserequest.setCvdInfo cvd_indicator, cvd_value
		track2purchaserequest.setCommcardInvoice commcard_invoice
		track2purchaserequest.setCommcardTaxAmount commcard_tax_amount

	out.setRequest track2purchaserequest.formatRequest(orderid, Amount, track2, CardNum, Expdate, pos_code)
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
	Response.Write "Avs Result:  " & out.getAvsResultCode & "<br>"
	Response.Write "Cvd Result:  " & out.getCvdResultCode & "<br>"
	Response.Write "dumpXML:  " & out.dumpXMLResponse & "<br>"
%>

</BODY>
</HTML>

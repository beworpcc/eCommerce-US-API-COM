<% @Language = "VBScript" %>
<% Response.buffer = true %>
<%
	store_id = "monusqa002"
	api_token = "qatoken"

	order_id = "order_id"
	pan="4242424242424242"
	exp_date = "1111"
	cust_id = "cust_id"
	crypt_type = "7"

	avs_street_number="123"
	avs_street_name="Main St."
	avs_zipcode="22102"

	Set out = server.CreateObject("Moneris.USRequest")
	out.initRequest store_id, api_token, "https://esplusqa.moneris.com/gateway_us/servlet/MpgRequest"

	Set recurupdaterequest = server.CreateObject("Moneris.USRecurUpdate")

		recurupdaterequest.setCustId cust_id
		recurupdaterequest.setOrderId order_id
		recurupdaterequest.setPan pan
		recurupdaterequest.setExpdate exp_date
		recurupdaterequest.setAvsInfo avs_street_number, avs_street_name, avs_zipcode
		recurupdaterequest.setRecurAmount "1.00"
		recurupdaterequest.setAddNumRecurs "20"
		recurupdaterequest.setTotalNumRecurs "999"
		recurupdaterequest.setHold "false"
		recurupdaterequest.setTerminate "false"
		recurupdaterequest.setPAccountNumber "Account a12345678 9876543"
		recurupdaterequest.setPresentationType "X"

	out.setRequest recurupdaterequest.formatRequest
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

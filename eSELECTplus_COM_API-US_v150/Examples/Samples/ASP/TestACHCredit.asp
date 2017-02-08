<% @Language = "VBScript" %>
<% Response.buffer = true %>
<%
	store_id = "monusqa024"
	api_token = "qatoken"

	order_id = "achCredit-"&Day(Date)&Month(Date)&Year(Date)&"-"&Hour(Now)&Minute(Now)&Second(Now)
	amount="5.00"

	'===========
	'ACH Info
	sec="ppd"
	cust_first_name="Joe"
	cust_last_name="Jones"
	cust_address1="123 Main St."
	cust_address2=""
	cust_city="Destine"
	cust_state="TX"
	cust_zip="32542"
	routing_num="123123122"
	account_num="123456782"
	check_num="112"
	account_type="checking"

	Set out = server.CreateObject("Moneris.USRequest")
	out.initRequest store_id, api_token, "https://esplusqa.moneris.com/gateway_us/servlet/MpgRequest"

	Set achcreditrequest = server.CreateObject("Moneris.USAchCredit")
		achcreditrequest.setAchInfo sec, cust_first_name, cust_last_name, cust_address1, cust_address2, cust_city, cust_state, cust_zip, routing_num, account_num, check_num, account_type

	out.setRequest achcreditrequest.formatRequest(order_id, amount)
	'out.setStatusCheck "True"
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
	'Response.Write "Status Code:  " & out.getStatusCode & "<br>"
	'Response.Write "Status Message:  " & out.getStatusMsg & "<br>"
	Response.Write "dumpXML:  " & out.dumpXMLResponse & "<br>"
%>

</BODY>
</HTML>

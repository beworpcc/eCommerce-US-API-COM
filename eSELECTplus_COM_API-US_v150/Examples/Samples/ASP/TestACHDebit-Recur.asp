<% @Language = "VBScript" %>
<% Response.buffer = true %>
<%
	store_id = "monusqa024"
	api_token = "qatoken"

	order_id = "SL-"&Day(Date)&Month(Date)&Year(Date)&"-"&Hour(Now)&Minute(Now)&Second(Now)
	amount = "1.00"
	customer_id = "Customer_identifier"

	Set out = server.CreateObject("Moneris.USRequest")
	out.initRequest store_id , api_token , "https://esplusqa.moneris.com/gateway_us/servlet/MpgRequest"

	Set achDebit = server.CreateObject("Moneris.USAchDebit")

	'Ach Info fields
	'The values can be a string, "" or nullvalue

	sec = "ppd"
	cust_first_name = "Bob"
	cust_last_name = "Smith"
	cust_address1 = "101 Main St"
	cust_address2 = "Apt 102"
	cust_city = "Chicago"
	cust_state = "IL"
	cust_zip = "123456"
	routing_num = "54321"
	account_num = "23456"
	check_num = "100"
	account_type = "savings"
	achDebit.setAchInfo sec, cust_first_name, cust_last_name, cust_address1, cust_address2, cust_city, cust_state, cust_zip, routing_num, account_num, check_num, account_type

	'================
	'Recurring setup.
	recur_type = "day"
	start_now = "true"
	start_date = "2007/08/02"
	num_recurs = "4"
	period = "1"
	recur_amount = "14.00"
	achDebit.setRecur recur_type, start_now, start_date, num_recurs, period, recur_amount

	achDebit.setCustId( customer_id )

	out.setRequest achDebit.formatRequest( order_id, amount )
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

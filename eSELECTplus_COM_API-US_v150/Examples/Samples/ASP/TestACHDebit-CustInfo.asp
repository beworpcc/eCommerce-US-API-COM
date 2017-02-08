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


	'Setting Customer Info data'
	first_name = "MyFirst_Name"
	last_name = "MyLast_Name"
	company_name = "My Company name"
	address = "My Address"
	city = "MyCity"
	province = "MyProvince"
	postal_code = "A3A B5B"
	country = "MyCountry"
	phone_number = "416-333-5555"
	fax = "416-333-6666"
	tax1 = "0.07"
	tax2 = "0.08"
	tax3 = "1.00"
	shipping_cost = "6.50"
	email = "q@q.com"
	instruction = "Make it so!"
	ShippingAddr = achDebit.formatAddress( first_name, last_name, company_name, address, city, province, postal_code, country, phone_number, fax, tax1, tax2, tax3, shipping_cost )
	BillingAddr = ShippingAddr

	item_name = "Hammer"
	item_quantity = "1"
	product_code = "hmr-001"
	item_amount = "6.99"
	achDebit.setItem item_name, item_quantity, product_code, item_amount

	achDebit.setCustInfo BillingAddr, ShippingAddr, email, instruction

	'Set optional Customer ID'
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

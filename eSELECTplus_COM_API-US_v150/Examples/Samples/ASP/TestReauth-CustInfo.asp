<% @Language = "VBScript" %>
<% Response.buffer = true %>
<%
	store_id = "monusqa002"
	api_token = "qatoken"

	order_id = "reauth-"&Day(Date)&Month(Date)&Year(Date)&"-"&Hour(Now)&Minute(Now)&Second(Now)
	orig_order_id="original_order_id"
	amount = "10.00"
	txn_number = "564788-0_10"
	crypt_type = "7"
	cust_id = "cust_id"

	Set out = server.CreateObject("Moneris.USRequest")
	out.initRequest store_id, api_token, "https://esplusqa.moneris.com/gateway_us/servlet/MpgRequest"

	Set reauth = server.CreateObject("Moneris.USReauth")
		reauth.setCustId cust_id
		billing=reauth.formatAddress("bfname", "blname", "bCompany Name", "baddress", "bcity", "bprovince", "bpostal", "bcountry", "bphone1", "bfax" , "btax1", "btax2", "btax3", "bshipping_cost")
		shipping=reauth.formatAddress("sfname", "slname", "sCompany Name", "saddress", "scity", "sprovince", "spostal", "scountry", "sphone1", "sfax" , "stax1", "stax2", "stax3", "sshipping_cost")
		reauth.setCustInfo billing,shipping,"aa@email.com","take this instruction"

	out.setRequest reauth.formatRequest(order_id, orig_order_id,txn_number,amount,crypt_type)
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
%>

</BODY>
</HTML>

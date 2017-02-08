<%
function getMDComponent( data, param_pos )
	mark_pos = 0
	'Get last character before component.
	for i = 1 to param_pos - 1
		mark_pos = instr( mark_pos + 1, data, ";" )
	next
	'Get number of characters for component.
	num_of_char = instr( mark_pos + 1, data, ";" ) - mark_pos
	'Return component'
	getMDComponent = mid( data, mark_pos + 1, num_of_char - 1 )
end function

Dim accept, useragent
Dim cavv
Dim mpiSucc
Dim order_id, store_id, api_token, charge_total, pan, expiry_date, md

if Request.Form( "PaRes" ) = "" then
	if Request.Form( "charge_total" ) = "" then
%>
		<!-- Initial page to collect payment info from customer. -->
		<html>
		<head><title>VBV Test Transaction</title>
		</head>
		<body>
		<h3>VBV Transaction Test Page</h3>
		<form action="http://localhost/API/US/transactions/TestPurchase-VBV.asp" method="post">
		<table>
			<tr>
				<td align="right">Order ID:</td>
				<td><input type=text name="order_id" value="pur00001"></td>
			</tr>
			<tr>
				<td align="right">Store ID:</td>
				<td><input type=text name="store_id" value="monusqa002"></td>
			</tr>
			<tr>
				<td align="right">API Token:</td>
				<td><input type=text name="api_token" value="qatoken"></td>
			</tr>
			<tr>
				<td align="right">Credit Card Number:</td>
				<td><input type="text" name="pan" value="4242424242424242"></td>
			</tr>
			<tr>
				<td align="right">Expiry Date:</td>
				<td><input type="text" name="expiry_date" value="1111"></td>
			</tr>
			<tr>
				<td align="right">Charge Total:</td>
				<td><input type="text" name="charge_total" value="99.00"></td>
			</tr>
			<tr>
				<td><input type="submit" value="Submit"></td>
			</tr>
		</table>
		</form>
		</body>
		</html>
<%
	else
		'Check for Visa transaction here and proceed if true.
		'Txn Transaction with MPI.
		order_id = Request.Form( "order_id" )
		store_id = Request.Form( "store_id" )
		api_token = Request.Form( "api_token" )
		charge_total = Request.Form( "charge_total" )
		pan = Request.Form( "pan" )
		expiry_date = Request.Form( "expiry_date" )
		md = order_id & ";"  & store_id & ";" & api_token & ";" & charge_total & ";" & pan & ";" & expiry_date & ";"

		Set out = Server.CreateObject( "Moneris.USRequest" )
		out.initMpiRequest "https://esplusqa.moneris.com/mpi/servlet/MpiServlet"

		accept = Request.ServerVariables( "HTTP_ACCEPT" )
		useragent = Request.ServerVariables( "HTTP_USER_AGENT" )

		Set purreq = Server.CreateObject( "Moneris.USMPIReq" )

		'Validate Order ID.
		s = len( order_id )
		if s <= 20 AND s > 0 then
			if s < 20 then
				order_id = order_id & "-"

				d = 20 - len( order_id )
				for i = 1 to d
					order_id = order_id & "0"
				next
			end if

			out.setRequest  purreq.formatRequest( store_id, api_token, purreq.formatTxnRequest( order_id, charge_total, pan, expiry_date, md, "http://localhost/API/US/transactions/TestPurchase-VBV.asp", accept, useragent ) )
			out.sendRequest
			mpiSucc = out.getMPISuccess

			if mpiSucc = "true" then
				Response.Write out.getMPIInlineForm	'Creates VBV PIN prompt on customer's browser.
			else
				'Check eSELECTplus Integration Guide documentation to determine how to proceed with non-VBV transaction.
				Response.write "Popup Window getMPISuccess NOT true <br><br>"	'Debug code.
				Response.Write out.dumpXMLResponse	'Debug code.
			end if
		else
			Response.write "Order ID must be between 1 to 20 characters in length!"
		end if
	end if
else
%>
	<html>
	<head>
	</head>
	<body>
<%
	'Acs Transaction with MPI.
	md = Request.Form( "MD" )
	order_id = getMDComponent( md, 1 )
	store_id = getMDComponent( md, 2 )
	api_token = getMDComponent( md, 3 )
	charge_total = getMDComponent( md, 4 )
	pan = getMDComponent( md, 5 )
	expiry_date = getMDComponent( md, 6 )

	Set out = Server.CreateObject( "Moneris.USRequest" )
	out.initMpiRequest "https://esplusqa.moneris.com/mpi/servlet/MpiServlet"

	Set purreq = Server.CreateObject( "Moneris.USMPIReq" )

	out.setRequest purreq.formatRequest( store_id, api_token, purreq.formatAcsRequest( Request.Form( "PaRes" ), md ) )

	out.sendRequest
	mpiSucc = out.getMPISuccess

	if mpiSucc = "true" then
		'Financial transaction with MPG.
		cavv = out.getMPICavv

		Set out = Server.CreateObject( "Moneris.USRequest" )
		out.initRequest store_id, api_token, "https://esplusqa.moneris.com/gateway_us/servlet/MpgRequest"

		Set purreq = Server.CreateObject( "Moneris.USCAVVPurchase" )

		'Recurring setup.
		'purreq.setRecur "week", "false", "2011/11/30", "3", "8", "13.50"

		out.setRequest purreq.formatRequest( order_id, charge_total, pan, expiry_date, cavv )
		out.sendRequest

		'Display financial transaction result.
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
		Response.Write "RecurSuccess:  " & out.getRecurSuccess & "<br>"
	else
		'Authentication of cardholder by financial institution failed,
		'check eSELECTplus Integration Guide documentation.
		Response.write "ACS getMPISuccess NOT true <br><br>"	'Debug code.
		Response.Write out.dumpXMLResponse	'Debug code.
	end if
%>
	</body>
	</html>
<%
end if
%>

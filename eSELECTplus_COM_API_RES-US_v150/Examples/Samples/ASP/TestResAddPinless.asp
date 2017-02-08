<% @Language = "VBScript" %>
<% Response.buffer = true %>
<%
	store_id = "monusqa002"
	api_token = "qatoken"

	cust_id = "602" & Time()
	phone = "999-888-7777"
	email = "my@email.com"
	note = "hello world"
	pan = "4496270000164824"
	presentation_type = "W"

	expdate = "1212"
	p_account_number = "1234567890123456789012345"

	Set out = server.CreateObject("Moneris.USResolverRequest")
	out.initRequest store_id, api_token, "esplusqa.moneris.com"

	Set request1 = server.CreateObject("Moneris.USResAddPinless")

	out.setRequest request1.formatRequest(cust_id, phone, email, note, pan, presentation_type)
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
	Response.Write "TransDate:  " & out.getTransDate & "<br>"
	Response.Write "TransTime:  " & out.getTransTime & "<br>"
	Response.Write "Complete:  " & out.getComplete & "<br>"
	Response.Write "TimedOut:  " & out.getTimedOut & "<br>"

	Response.Write "ResSuccess:  " & out.getResSuccess & "<br>"
	Response.Write "PaymentType:  " & out.getPaymentType & "<br>"
	Response.Write "ResDataDataKey:  " & out.getDataKey & "<br>"
	Response.Write "ResDataPaymentType:  " & out.getResDataPaymentType & "<br>"
	Response.Write "ResDataCustId:  " & out.getResDataCustId & "<br>"
	Response.Write "ResDataPhone:  " & out.getResDataPhone & "<br>"
	Response.Write "ResDataEmail:  " & out.getResDataEmail & "<br>"
	Response.Write "ResDataNote:  " & out.getResDataNote & "<br>"
	Response.Write "ResDataMaskedPan:  " & out.getResDataMaskedPan & "<br>"
	Response.Write "ResDataExpDate:  " & out.getResDataExpDate & "<br>"
	Response.Write "Presentation Type:  " & out.getResDataPresentationType & "<br>"
	Response.Write "Account Number:  " & out.getResDataPAccountNumber & "<br>"
	Response.Write "ResDataCryptType:  " & out.getResDataCryptType & "<br>"
	Response.Write "dumpXML:  " & out.dumpXMLResponse & "<br>"
%>

</BODY>
</HTML>

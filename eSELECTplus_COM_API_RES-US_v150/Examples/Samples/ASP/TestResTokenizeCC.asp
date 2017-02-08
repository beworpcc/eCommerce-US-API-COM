<% @Language = "VBScript" %>
<% Response.buffer = true %>
<%
	store_id = "monusqa002"
	api_token = "qatoken"

	order_id = "602" & Time()
	txn_number = "123456-0_1"
	phone = "1-777-888-9999"
	email = "i.have.no@email.com"
	note = "how are you"
	cust_id="cus_id"

	street_num = "123"
	street_name = "Main Street"
	zipcode = "90210"

	Set out = server.CreateObject("Moneris.USResolverRequest")
	out.initRequest store_id, api_token, "esplusqa.moneris.com"

	Set request1 = server.CreateObject("Moneris.USResTokenizeCC")
	request1.setCustId cust_id
	request1.setPhone phone
	request1.setNote note
	request1.setEmail email
	request1.setAvsInfo street_num, street_name, zipcode

	out.setRequest request1.formatRequest(order_id, txn_number)
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
	Response.Write "Sec:  " & out.getResDataSec & "<br>"
	Response.Write "Cust First Name:  " & out.getResDataCustFirstName & "<br>"
	Response.Write "Cust Last Name:  " & out.getResDataCustLastName & "<br>"
	Response.Write "Cust Address 1:  " & out.getResDataCustAddress1 & "<br>"
	Response.Write "Cust Address 2:  " & out.getResDataCustAddress2 & "<br>"
	Response.Write "Cust City:  " & out.getResDataCustCity & "<br>"
	Response.Write "Cust State:  " & out.getResDataCustState & "<br>"
	Response.Write "Cust Zip:  " & out.getResDataCustZip & "<br>"
	Response.Write "Routing Num:  " & out.getResDataRoutingNum & "<br>"
	Response.Write "Masked Account Num:  " & out.getResDataMaskedAccountNum & "<br>"
	Response.Write "Check Num:  " & out.getResDataCheckNum & "<br>"
	Response.Write "Account Type:  " & out.getResDataAccountType & "<br>"
	Response.Write "dumpXML:  " & out.dumpXMLResponse & "<br>"
%>

</BODY>
</HTML>

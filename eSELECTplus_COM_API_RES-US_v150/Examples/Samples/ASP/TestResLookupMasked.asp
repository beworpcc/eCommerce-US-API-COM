<% @Language = "VBScript" %>
<% Response.buffer = true %>
<%
	store_id = "monusqa002"
	api_token = "qatoken"

	order_id = "602" & Time()
	data_key = "Q6w3v99jgKs2snTocQmJjFxGc"

	Set out = server.CreateObject("Moneris.USResolverRequest")
	out.initRequest store_id, api_token, "esplusqa.moneris.com"

	Set request1 = server.CreateObject("Moneris.USResLookupMasked")

	out.setRequest request1.formatRequest(data_key)
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
	Response.Write "Cust Last Name:  " & out.getResDataLastName & "<br>"
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

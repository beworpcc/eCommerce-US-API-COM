<% @Language = "VBScript" %>
<% Response.buffer = true %>
<%
	store_id = "monusqa002"
	api_token = "qatoken"

	Set out = server.CreateObject("Moneris.USResolverRequest")
	out.initRequest store_id, api_token, "esplusqa.moneris.com"

	Set request1 = server.CreateObject("Moneris.USResGetExpiring")

	out.setRequest request1.formatRequest()
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
	Response.Write "ResSuccess:  " & out.getResSuccess & "<br>"
	Response.Write "PaymentType:  " & out.getPaymentType & "<br>"
	Response.Write "ResDataDataKey:  " & out.getResDataDataKey & "<br>"
	Response.Write "ResDataPaymentType:  " & out.getResDataPaymentType & "<br>"
	Response.Write "ResDataCustId:  " & out.getResDataCustId & "<br>"
	Response.Write "ResDataPhone:  " & out.getResDataPhone & "<br>"
	Response.Write "ResDataEmail:  " & out.getResDataEmail & "<br>"
	Response.Write "ResDataNote:  " & out.getResDataNote & "<br>"
	Response.Write "ResDataMaskedPan:  " & out.getResDataMaskedPan & "<br>"
	Response.Write "ResDataExpDate:  " & out.getResDataExpDate & "<br>"
	Response.Write "ResDataCryptType:  " & out.getResDataCryptType & "<br>"
	Response.Write "dumpXML:  " & out.dumpXMLResponse & "<br>"
%>

</BODY>
</HTML>

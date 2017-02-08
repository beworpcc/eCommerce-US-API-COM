<% @Language = "VBScript" %>
<% Response.buffer = true %>
<%
store_id = "monusqa002"
api_token = "qatoken"
order_id = "encTrack2ForcePost_" & DateDiff("s", CDate("1970/01/01 00:00:00"), Now)
custid = "Customer 001"

Amount = "10.00"
enc_track2 = ";5258968987035454=06061015454001060101?"
pos_code = "00"
device_type = "idtech"

Set out = server.CreateObject("Moneris.USRequest")
out.initRequest store_id, api_token, "https://esplusqa.moneris.com/gateway_us/servlet/MpgRequest"

Set encTrack2ForcePostRequest = server.CreateObject("Moneris.USEncTrack2ForcePost")
        encTrack2ForcePostRequest.setCustId custid

out.setRequest encTrack2ForcePostRequest.formatRequest(order_id, Amount, enc_track2, pos_code, device_type)
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
Response.Write "Card Type:  " & out.getCardType & "<br>"
Response.Write "Reference Number:  " & out.getReferenceNum & "<br>"
Response.Write "Transaction ID:  " & out.getTransID & "<br>"
Response.Write "Auth Code:  " & out.getAuthCode & "<br>"
Response.Write "Transaction Time:  " & out.getTransTime & "<br>"
Response.Write "Transaction Date:  " & out.getTransDate & "<br>"
Response.Write "Complete:  " & out.getCompleteStatus & "<br>"
Response.Write "Timeout:  " & out.getTimedoutStatus & "<br>"
Response.Write "Ticket:  " & out.getTicket & "<br>"
Response.Write "MaskedPan:  " & out.getMaskedPan & "<br>"
Response.Write "dumpXML:  " & out.dumpXMLResponse & "<br>"
%>

</BODY>
</HTML>

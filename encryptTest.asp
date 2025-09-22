<% @CODEPAGE="65001" language="VBScript" %>
<%
Response.CharSet="utf-8"
Session.codepage="65001"
Response.codepage="65001"
Response.ContentType="text/html;charset=utf-8"
%>
<%

Dim objCrypto : Set objCrypto = Server.CreateObject("SBCryptoUtil.CryptoUtil.1") '암복호화 객체 생성(AES-256-ECB -> Base64)


Dim AES256_KEY : AES256_KEY = "pgSettle30y739r82jtd709yOfZ2yK5K" '암호화 키값
Dim plainText : plainText ="TEST" '평문
Dim cipherText '암호문


cipherText = objCrypto.encryptBase64(AES256_KEY, plainText) '암호화
response.write("PlainText["&plainText&"] ---> CipherText["&cipherText&"]<br>")

plainText = objCrypto.decryptBase64(AES256_KEY, cipherText) '복호화
response.write("CipherText["&cipherText&"] ---> PlainText["&plainText&"]<br>")


set objCrypto = nothing
%>

<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>S'Pay</title>
<style type="text/css">
#STPG_RSLT		{font-family:굴림; font-size:10pt;}
#STPG_RSLT h4	{background-color:#f1f1f1;padding:4px;margin:2px;}
</style>
</head>
<body>
<h3>AES256 암복호화 TEST</h3>
<div id="STPG_RSLT"> 
    <input type="button" value="돌아가기" style="margin-top:20px;" onclick="location.href='pay_form.asp'">
</div>
</body>
</html>
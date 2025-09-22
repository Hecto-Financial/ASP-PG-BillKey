<% @CODEPAGE="65001" language="VBScript" %>
<% option explicit %>
<% response.Charset = "UTF-8" %>
<!--#include virtual="/npg/inc/KISA_SHA256.asp"-->
<!--#include virtual="/npg/inc/settleUtils.asp"-->
<!--#include virtual="/npg/inc/config.asp"-->
<%
'============================================================================================================================================
'   요청 파라미터
'============================================================================================================================================
Dim REQ_PARAMS : set REQ_PARAMS = JSON.parse("{ ""params"":{}, ""data"":{} }")
REQ_PARAMS.params.set "mchtId",     "" & request.Form("mchtId")     '상점아이디
REQ_PARAMS.params.set "ver",        "" & request.Form("ver")        '버전
REQ_PARAMS.params.set "method",     "" & request.Form("method")     '결제수단
REQ_PARAMS.params.set "bizType",    "" & request.Form("bizType")    '업무구분
REQ_PARAMS.params.set "encCd",      "" & request.Form("encCd")      '암호화구분
REQ_PARAMS.params.set "mchtTrdNo",  "" & request.Form("mchtTrdNo")  '상점주문번호
REQ_PARAMS.params.set "trdDt",      "" & request.Form("trdDt")      '요청일자
REQ_PARAMS.params.set "trdTm",      "" & request.Form("trdTm")      '요청시간
REQ_PARAMS.params.set "mobileYn",   "" & request.Form("mobileYn")   '모바일여부
REQ_PARAMS.params.set "osType",     "" & request.Form("osType")     '운영체제 구분

REQ_PARAMS.data.set "pktHash"   , "" & request.Form("pktHash")      '해쉬값
REQ_PARAMS.data.set "cardNo"    , "" & request.Form("cardNo")       '카드번호
REQ_PARAMS.data.set "idntNo"    , "" & request.Form("idntNo")       '식별번호
REQ_PARAMS.data.set "vldDtMon"  , "" & request.Form("vldDtMon")     '유효기간(월)
REQ_PARAMS.data.set "vldDtYear" , "" & request.Form("vldDtYear")    '유효기간(년)
REQ_PARAMS.data.set "cardPwd"   , "" & request.Form("cardPwd")      '카드비밀번호
REQ_PARAMS.data.set "mchtCustNm", "" & request.Form("mchtCustNm")   '고객명
REQ_PARAMS.data.set "mchtCustId", "" & request.Form("mchtCustId")   '고객아이디
REQ_PARAMS.data.set "keyRegYn"  , "" & request.Form("keyRegYn")     '빌키발급요청여부


'============================================================================================================================================
'   응답 파라미터 선언
'============================================================================================================================================
Dim RES_PARAMS : set RES_PARAMS = JSON.parse("{ ""params"":{}, ""data"":{} }")
RES_PARAMS.params.set "mchtId", ""      '상점아이디
RES_PARAMS.params.set "ver", ""         '버전
RES_PARAMS.params.set "method", ""      '결제수단
RES_PARAMS.params.set "bizType", ""     '업무구분
RES_PARAMS.params.set "encCd", ""       '암호화구분
RES_PARAMS.params.set "mchtTrdNo", ""   '상점주문번호
RES_PARAMS.params.set "trdNo", ""       '헥토파이낸셜거래번호
RES_PARAMS.params.set "trdDt", ""       '거래일자
RES_PARAMS.params.set "trdTm", ""       '거래시간
RES_PARAMS.params.set "outStatCd", ""   '거래상태코드
RES_PARAMS.params.set "outRsltCd", ""   '결과코드
RES_PARAMS.params.set "outRsltMsg", ""  '결과메세지

RES_PARAMS.data.set "pktHash", ""       '해쉬값
RES_PARAMS.data.set "cardNo", ""        '카드번호
RES_PARAMS.data.set "issrId", ""        '발급사아이디
RES_PARAMS.data.set "cardNm", ""        '카드사명
RES_PARAMS.data.set "cardKind", ""      '카드종류명
RES_PARAMS.data.set "billKey", ""       '빌키


'AES256 암호화 필요한 요청파라미터
Dim ENCRYPT_PARAMS : ENCRYPT_PARAMS = array("cardNo", "idntNo", "vldDtMon", "vldDtYear", "cardPwd")

'AES256 복호화 필요한 응답파라미터
Dim DECRYPT_PARAMS : DECRYPT_PARAMS = array()


'============================================================================================================================================
'   SHA256 해쉬 처리
'조합 필드 : 요청일자 + 요청시간 + 상점아이디 + 상점주문번호 + "0" + 라이센스키
'============================================================================================================================================
Dim hashPlain : hashPlain = ""
Dim hashCipher : hashCipher = ""
hashPlain = REQ_PARAMS.params.get("trdDt") &_
            REQ_PARAMS.params.get("trdTm") &_
            REQ_PARAMS.params.get("mchtId") &_
            REQ_PARAMS.params.get("mchtTrdNo") &_
            "0" &_
            LICENSE_KEY
On error resume next
hashCipher = SHA256_Encrypt(hashPlain)'해쉬 값 계산
If Err.Number <> 0 Then
    call log_message(LOG_FILE, "[" & REQ_PARAMS.params.get("mchtTrdNo") & "][SHA256 HASHING] Hashing Fail! " & err.number & " : " & err.description  )
else
    call log_message(LOG_FILE, "[" & REQ_PARAMS.params.get("mchtTrdNo") & "][SHA256 HASHING] Plain Text[" & hashPlain & "] ---> Cipher Text[" & hashCipher & "]")
    REQ_PARAMS.data.set "pktHash", hashCipher 'SHA256 해쉬 결과 저장
End If
On Error GoTo 0



'============================================================================================================================================
'   AES256 암호화 처리
'============================================================================================================================================
On error resume next
Dim objCrypto : Set objCrypto = Server.CreateObject("SBCryptoUtil.CryptoUtil.1") '암복호화 객체 생성
Dim aesPlain : aesPlain = ""
Dim aesCipher : aesCipher = ""
Dim i
for each i in ENCRYPT_PARAMS
    aesPlain = REQ_PARAMS.data.get(i)
    if "" <> aesPlain then
        aesCipher = objCrypto.EncryptBase64(AES256_KEY, aesPlain) 'AES256 암호화
        REQ_PARAMS.data.set i, aesCipher '암호화 결과 값 세팅
        call log_message(LOG_FILE, "[" & REQ_PARAMS.params.get("mchtTrdNo") & "][AES256 Encrypt] " & i & "[" & aesPlain & "] ---> [" & aesCipher & "]")
    end if
next

If Err.Number <> 0 Then
    call log_message(LOG_FILE, "[" & REQ_PARAMS.params.get("mchtTrdNo") & "][AES256 Encrypt] Fail! " & err.number & " : " & err.description  )
End If
On Error GoTo 0



'============================================================================================================================================
'   API URL 설정
'============================================================================================================================================
Dim requestUrl : requestUrl = SERVER_URL & "/spay/APICardAuth.do"
call log_message(LOG_FILE, "[" & REQ_PARAMS.params.get("mchtTrdNo") & "][SEND POST] Request url :  " & requestUrl  )



'============================================================================================================================================
'   API호출(가맹점->헥토파이낸셜) 및 응답 처리
'============================================================================================================================================
On error resume next
call log_message(LOG_FILE, "[" & REQ_PARAMS.params.get("mchtTrdNo") & "][SEND POST] Request :  " & JSON.stringify(REQ_PARAMS)  )
Dim resData : set resData = sendPost(requestUrl, JSON.stringify(REQ_PARAMS) , CONN_TIMEOUT, READ_TIMEOUT)
If Err.Number <> 0 Then
    call log_message(LOG_FILE, "[" & REQ_PARAMS.params.get("mchtTrdNo") & "][SEND POST] Fail! " & err.number & " : " & err.description  )
Else
    call log_message(LOG_FILE, "[" & REQ_PARAMS.params.get("mchtTrdNo") & "][SEND POST] Response : " &  JSON.stringify(resData) )
End If
On Error GoTo 0


'============================================================================================================================================
'   응답 파라미터 세팅
'============================================================================================================================================
if resData.get("params") <> "" then
    RES_PARAMS.params.set "mchtId", resData.params.get("mchtId")            '상점아이디
    RES_PARAMS.params.set "ver", resData.params.get("ver")                  '버전
    RES_PARAMS.params.set "method", resData.params.get("method")            '결제수단
    RES_PARAMS.params.set "bizType", resData.params.get("bizType")          '업무구분
    RES_PARAMS.params.set "encCd", resData.params.get("encCd")              '암호화구분
    RES_PARAMS.params.set "mchtTrdNo", resData.params.get("mchtTrdNo")      '상점주문번호
    RES_PARAMS.params.set "trdNo", resData.params.get("trdNo")              '헥토파이낸셜거래번호
    RES_PARAMS.params.set "trdDt", resData.params.get("trdDt")              '거래일자
    RES_PARAMS.params.set "trdTm", resData.params.get("trdTm")              '거래시간
    RES_PARAMS.params.set "outStatCd", resData.params.get("outStatCd")      '거래상태코드
    RES_PARAMS.params.set "outRsltCd", resData.params.get("outRsltCd")      '결과코드
    RES_PARAMS.params.set "outRsltMsg", resData.params.get("outRsltMsg")    '결과메세지
end if
if resData.get("data") <> "" then
    RES_PARAMS.data.set "pktHash", resData.data.get("pktHash")              '해쉬값
    RES_PARAMS.data.set "cardNo", resData.data.get("cardNo")                '카드번호
    RES_PARAMS.data.set "issrId", resData.data.get("issrId")                '발급사아이디
    RES_PARAMS.data.set "cardNm", resData.data.get("cardNm")                '카드사명
    RES_PARAMS.data.set "cardKind", resData.data.get("cardKind")            '카드종류명
    RES_PARAMS.data.set "billKey", resData.data.get("billKey")              '빌키
end if
'============================================================================================================================================
'   AES256 복호화 처리
'============================================================================================================================================
On error resume next
aesPlain = ""
aesCipher = ""
for each i in DECRYPT_PARAMS
    aesCipher = Trim(RES_PARAMS.data.get(i))
    if "" <> aesCipher then
        aesPlain = objCrypto.DecryptBase64(AES256_KEY, aesCipher) 'AES256 복호화
        RES_PARAMS.data.set i, aesPlain '복호화 결과 값 세팅
        call log_message(LOG_FILE, "[" & REQ_PARAMS.params.get("mchtTrdNo") & "][AES256 Decrypt] " & i & "[" & aesCipher & "] ---> [" & aesPlain & "]")
    end if
next

If Err.Number <> 0 Then
    call log_message(LOG_FILE, "[" & REQ_PARAMS.params.get("mchtTrdNo") & "][AES256 Decrypt] fail! " & err.number & " : " & err.description  )
End If
On Error GoTo 0
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
<h3>응답 결과</h3>
<div id="STPG_RSLT"> 
    <table>
     	<tr>
            <td colspan="2" style="text-align: center;"><h4>params</h4></td>
        </tr>
        <tr>
            <td>mchtId[상점아이디]</td>
            <td><%= RES_PARAMS.params.get("mchtId") %></td>
        </tr>
        <tr>
            <td>ver[버전]</td>
            <td><%= RES_PARAMS.params.get("ver") %></td>
        </tr>
        <tr>
            <td>method[결제수단]</td>
            <td><%= RES_PARAMS.params.get("method") %></td>
        </tr>
        <tr>
            <td>bizType[업무구분]</td>
            <td><%= RES_PARAMS.params.get("bizType") %></td>
        </tr>
        <tr>
            <td>encCd[암호화구분]</td>
            <td><%= RES_PARAMS.params.get("encCd") %></td>
        </tr>
        <tr>
            <td>mchtTrdNo[상점주문번호]</td>
            <td><%= RES_PARAMS.params.get("mchtTrdNo") %></td>
        </tr>
        <tr>
            <td>trdNo[헥토파이낸셜 거래번호]</td>
            <td><%= RES_PARAMS.params.get("trdNo") %></td>
        </tr>
        <tr>
            <td>trdDt[거래일자]</td>
            <td><%= RES_PARAMS.params.get("trdDt") %></td>
        </tr>
        <tr>
            <td>trdTm[거래시간]</td>
            <td><%= RES_PARAMS.params.get("trdTm") %></td>
        </tr>
        <tr>
            <td>outStatCd[거래상태코드]</td>
            <td><%= RES_PARAMS.params.get("outStatCd") %></td>
        </tr>
        <tr>
            <td>outRsltCd[거래결과코드]</td>
            <td><%= RES_PARAMS.params.get("outRsltCd") %></td>
        </tr>
        <tr>
            <td>outRsltMsg[결과메세지]</td>
            <td><%= RES_PARAMS.params.get("outRsltMsg") %></td>
        </tr>
     	<tr>
            <td colspan="2" style="text-align: center;"><h4>data</h4></td>
        </tr>
        <tr>
            <td>pktHash[해쉬값]</td>
            <td><%= RES_PARAMS.data.get("pktHash") %></td>
        </tr>
        <tr>
            <td>cardNo[카드번호]</td>
            <td><%= RES_PARAMS.data.get("cardNo") %></td>
        </tr>
        <tr>
            <td>issrId[발급사아이디]</td>
            <td><%= RES_PARAMS.data.get("issrId") %></td>
        </tr>
        <tr>
            <td>cardNm[카드사명]</td>
            <td><%= RES_PARAMS.data.get("cardNm") %></td>
        </tr>
        <tr>
            <td>cardKind[카드종류명]</td>
            <td><%= RES_PARAMS.data.get("cardKind") %></td>
        </tr>
        <tr style="background-color:yellow;">
            <td>billKey[빌키]</td>
            <td><%= RES_PARAMS.data.get("billKey") %></td>
        </tr>

        <tr>
            <td colspan="2" style="text-align: center;">
            <input type="button" value="돌아가기" style="margin-top:20px;" onclick="location.href='authAPI_form.asp'">
            </td>
        </tr>
    </table>
</div>
</body>
</html>
<%
set REQ_PARAMS = nothing
set RES_PARAMS = nothing
set resData = nothing
set objCrypto = nothing
%>
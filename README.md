# ASP-PG-BillKey

헥토파이낸셜 신용카드 비/구인증(빌키) API 연동을 위한 ASP 샘플 코드입니다.

## 📋 개요

본 샘플코드는 **API 직접 호출(Non-UI) 방식**이며, 결제창(UI) 방식이 아닙니다.
- 결제창(UI) 방식을 원하실 경우, 표준결제창 연동하시면 됩니다.
- 1회차는 Non-UI 또는 UI 방식으로 취사선택하여 결제하시면 됩니다.
- 2회차 결제는 발급받으신 빌키로 API 직접 호출하여 결제하시면 됩니다.

**⚠️ 주의사항**: ASP 이용 가맹점의 경우 DLL 설치가이드(ASP Classic)를 참조하여 설치 바랍니다.

## 파일 구조

```
/(Project Root)
│
│     index.html			<--- index페이지
│     encryptTest.asp		<--- COM+ 컴포넌트 호출 테스트 페이지(암복호화 모듈)
│  
│     pay_form.asp			<--- 결제 API 양식
│     billKey_form.asp		<--- 빌키 결제 API 양식
│     pay_showResult.asp		<--- 결제 처리 및 결과 화면
│  
│     authAPI_form.asp		<--- 빌키 발급 API 양식
│     authAPI_showResult.asp		<--- 빌키 발급 및 결과 출력
│
│     cancel_form.asp		<--- 취소 API 양식
│     cancel_showResult.asp		<--- 취소 처리 및 결과 화면
│  
│     receiveNoti.asp		<--- 결제 완료 후 노티 수신 페이지
│     processNoti.asp		<--- 노티 수신 후 처리하는 페이지
│  
│  
└─npg
│     config.asp			<--- 기본정보 설정파일(*자사에 맞게 변경 필요)
│     json2.asp			<--- JSON 라이브러리
│     KISA_SHA256.asp		<--- KISA에서 배포한 SHA256라이브러리
│     settleUtils.asp			<--- 헥토파이낸셜 유틸 라이브러리
│
└─DLL
        libiconv.dll			<--- SBCryptoUtil.dll 에서 의존하는 iconv dll
        libiconvD.dll			<--- SBCryptoUtil.dll 에서 의존하는 iconv dll
        SBCryptoUtil.dll		<--- AES256 암호화 동적 라이브러리
```

## 파일 설명

### 공통 페이지
- **index.html**: 인덱스 페이지입니다.
- **encryptTest.asp**: 암복호화 컴포넌트 호출 테스트 페이지입니다.
- **config.asp**: 상점아이디, 암복호화키 등을 설정할 수 있는 설정파일입니다.
- **receiveNoti.asp**: 결제 또는 취소 처리가 완료된 후, 헥토파이낸셜에서 가맹점으로 전달하는 노티(결과통보)를 수신하는 페이지입니다.
- **processNoti.asp**: receiveNoti.asp에서 결제 또는 취소의 성공/실패에 따라 적절한 로직을 수행하는 메소드를 정의한 파일입니다.

### 결제 관련 페이지
- **pay_form.asp**: 결제 API 양식으로서, 빌키서비스상점의 경우 응답으로 빌키가 발급됩니다.
- **billKey_form.asp**: 발급받은 빌키로 결제하는 API 양식입니다.
- **pay_showResult.asp**: 헥토파이낸셜과 Server to Server 로 커넥션하여, 결제 요청을 하고 응답을 받아 결과를 출력하는 페이지입니다.
- **authAPI_form.asp**: 빌키 발급 API 양식으로서, 결제하지 않고 응답으로 빌키가 발급됩니다.
- **authAPI_showResult.asp**: 헥토파이낸셜과 Server to Server 로 커넥션하여, 요청을 하고 응답을 받아 결과를 출력하는 페이지입니다.

### 취소 관련 페이지
- **cancel_form.asp**: 취소 요청시 사용자로부터 정보를 입력받는 Form 페이지입니다.
- **cancel_showResult.asp**: 헥토파이낸셜과 Server to Server 로 커넥션하여, 취소 요청을 하고 응답을 받아 결과를 출력하는 페이지입니다.

## 프로세스 처리 순서

- **결제 API(빌키 발급 포함)**: pay_form.asp -> pay_showResult.asp
- **빌키 결제 API**: billKey_form.asp -> pay_showResult.asp
- **빌키 발급 API**: authAPI_form.asp -> authAPI_showResult.asp
- **취소 처리 순서**: cancel_form.asp -> cancel_showResult.asp
- **노티 처리 순서**: receiveNoti.asp -> processNoti.asp

## config.asp 설정파일 변수 설명

- **PG_MID**: 상점아이디. 테스트환경에서의 상점아이디는 샘플소스에 기재되어있습니다. 상용테스트시에는 헥토파이낸셜에서 발급한 MID로 설정하셔야 합니다. 이 값은 외부에 노출되어서는 안됩니다.
- **LICENSE_KEY**: MID당 하나의 라이센스키가 발급됩니다. SHA256해쉬체크 용도로 사용됩니다. 이 값은 외부에 노출되어서는 안됩니다.
- **AES256_KEY**: 개인정보/민감정보를 암복호화 하는데 사용되는 키로서, 외부에 노출되어서는 안됩니다.
- **SERVER_URL**: 헥토파이낸셜 결제 처리 서버의 URL입니다. 변경하지 마십시오.
- **CONN_TIMEOUT**: 헥토파이낸셜 API통신 연결 타임아웃입니다.
- **READ_TIMEOUT**: 헥토파이낸셜 API통신 수신 타임아웃입니다.

## 노티 수신 페이지

- **파일명**: receiveNoti.asp
- 결제 또는 취소 완료 후 헥토파이낸셜 서버에서 콜백으로 호출하게 되는 페이지이며, 헥토파이낸셜에서 가맹점으로 노티를 전송합니다.
- 노티 수신 페이지에서는 전송받은 노티 파라미터들을 적절히 사용하여 가맹점의 실제 내부데이터, DB를 처리하시면 됩니다.
- 헥토파이낸셜측에서 전송한 결제 결과(성공이든 실패든)에 상관없이 가맹점에서 성공적으로 내부데이터를 처리하셨다면 OK를, 아니라면 FAIL을 리턴하시면 됩니다.

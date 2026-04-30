# ASP-PG-BillKey

헥토파이낸셜 신용카드 비인증(빌키) API 연동을 위한 ASP(Classic ASP) 샘플 코드입니다.

> 자세한 연동 방법은 [헥토파이낸셜 개발자 센터](https://developers.hectofinancial.co.kr)를 참고하세요.

---

## 개요

본 샘플코드는 **API 직접 호출(Non-UI) 방식**이며, 결제창(UI) 방식이 아닙니다.

- 1회차 결제 시 Non-UI 또는 UI 방식 중 선택하여 결제하면 됩니다.
- 1회차 결제 응답으로 빌키가 발급됩니다.
- 2회차 이후 결제는 발급받은 빌키로 API를 직접 호출하여 결제합니다.
- 결제창(UI) 방식을 원하실 경우, 표준결제창 샘플을 사용하세요.

> **주의**: ASP 이용 가맹점의 경우 DLL 설치 가이드(ASP Classic)를 참조하여 DLL을 먼저 설치해야 합니다.

---

## 파일 구조

```
/(Project Root)
│  index.html                    # 인덱스 페이지
│  encryptTest.asp               # COM+ 컴포넌트 호출 테스트 페이지 (암복호화 모듈)
│
│  pay_form.asp                  # 결제 요청 폼 (빌키 발급 포함)
│  billKey_form.asp              # 빌키 결제 요청 폼
│  pay_showResult.asp            # 결제 처리 및 결과 화면
│
│  authAPI_form.asp              # 빌키 발급 전용 API 폼
│  authAPI_showResult.asp        # 빌키 발급 처리 및 결과 화면
│
│  cancel_form.asp               # 취소 요청 폼
│  cancel_showResult.asp         # 취소 처리 및 결과 화면
│
│  receiveNoti.asp               # 노티 수신 페이지
│  processNoti.asp               # 노티 처리 메소드 정의 파일
│
├─npg/inc/
│      config.asp                # 기본 정보 설정 파일 (상점 환경에 맞게 수정 필요)
│      json2.asp                 # JSON 라이브러리
│      KISA_SHA256.asp           # KISA 배포 SHA256 라이브러리
│      settleUtils.asp           # 헥토파이낸셜 유틸 라이브러리
│
└─DLL/                           # AES256 암호화 동적 라이브러리
       libiconv.dll
       libiconvD.dll
       SBCryptoUtil.dll
```

---

## 페이지 처리 순서

| 기능 | 순서 |
|------|------|
| 결제 API (빌키 발급 포함) | `pay_form.asp` → `pay_showResult.asp` |
| 빌키 결제 | `billKey_form.asp` → `pay_showResult.asp` |
| 빌키 발급 API | `authAPI_form.asp` → `authAPI_showResult.asp` |
| 취소 | `cancel_form.asp` → `cancel_showResult.asp` |
| 노티 수신 | `receiveNoti.asp` → `processNoti.asp` |

---

## 설정 (npg/inc/config.asp)

| 변수 | 설명 |
|------|------|
| `PG_MID` | 상점아이디. 테스트용 MID는 샘플에 포함되어 있으며, 운영 시 발급받은 MID로 교체하세요. **외부 노출 금지** |
| `LICENSE_KEY` | MID별 발급되는 라이센스키. SHA-256 해시 검증에 사용됩니다. **외부 노출 금지** |
| `AES256_KEY` | 개인정보/민감정보 AES-256 암복호화 키. **외부 노출 금지** |
| `SERVER_URL` | 헥토파이낸셜 처리 서버 URL. 테스트/운영 URL 주석 참고 후 변경하세요. |
| `CONN_TIMEOUT` | API 통신 연결 타임아웃 (ms) |
| `READ_TIMEOUT` | API 통신 수신 타임아웃 (ms) |
| `LOG_DIR` | 로그 파일 저장 디렉터리. 해당 디렉터리가 없으면 로그가 생성되지 않습니다. |
| `LOG_FILE` | 일반 거래 로그 파일명 |
| `NOTI_LOG_FILE` | 노티 관련 로그 파일명 |

---

## 참고

- 테스트 환경: `https://tbgw.settlebank.co.kr`
- 운영 환경: `https://gw.settlebank.co.kr`
- 개발자 센터: [https://developers.hectofinancial.co.kr](https://developers.hectofinancial.co.kr)

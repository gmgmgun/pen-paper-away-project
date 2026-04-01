# PPAP (Pen-Paper-Away-Project) 🖋️

```text
=============================================================================
██████╗ ██████╗  █████╗ ██████╗
██╔══██╗██╔══██╗██╔══██╗██╔══██╗
██████╔╝██████╔╝███████║██████╔╝
██╔═══╝ ██╔═══╝ ██╔══██║██╔═══╝
██║     ██║     ██║  ██║██║
╚═╝     ╚═╝     ╚═╝  ╚═╝╚═╝
=============================================================================
             🖋️ Pen-Paper-Away-Project: It's Not Pen Pineapple Apple Pen.
```

## 📝 Project Overview

PPAP는 기존의 번거로운 법인 차량 운행 수기 기록 방식을 디지털로 전환(DX)하여 업무 효율성을 극대화하는 프로젝트입니다. 종이와 펜(Pen & Paper)을 멀리하고(Away), 스마트한 데이터 관리를 지향합니다.

## 🚀 Why PPAP? (Problem Definition)

회사의 소중한 자산인 법인 차량을 관리함에 있어 기존 방식은 다음과 같은 Pain Point가 있었습니다.

- **불편한 기록:** 매번 차에 비치된 장부를 꺼내 수기로 작성해야 하는 번거로운 프로세스.
- **데이터 불일치:** 계기판 숫자를 잘못 적거나, 바쁜 일정으로 인해 기록을 누락하는 사례 발생.
- **관리 오버헤드:** 정산을 위해 수기 장부 데이터를 다시 엑셀로 옮기는 단순 반복 업무 발생.

## 🛠 Tech Stack & Architecture

가장 빠르고 비용 효율적인 기술 스택을 선택했습니다.

- **Entry:** QR Code (차량 대시보드 부착)
- **Input:** Google Apps Script Web App (Mobile Optimized HTML)
- **Database:** Google Sheets (`RAW_운행일지` 시트)
- **Logic:** Google Apps Script (GAS)
- **Config:** Script Properties (직원 목록, 차량 정보 등 런타임 주입)

## ✨ Key Features

- **Scan & Go:** QR 코드 스캔 한 번으로 즉시 기록 페이지 접속. URL 파라미터(`?car=차량번호`)로 차량 자동 인식.
- **Auto Calculation:** $Distance = Final - Initial$ 수식을 통해 주행 거리 자동 산출.
- **Cloud Sync:** 모든 데이터는 구글 시트(`RAW_운행일지`)에 실시간 저장되어 별도의 타이핑 작업 불필요.
- **Fixed User Support:** 고정 사용자 차량(전용 차량)은 성명 선택 없이 자동으로 운전자 정보 입력.
- **Business Trip Mode:** 출장용 차량은 방문 거래처 입력 UI가 별도로 활성화.
- **Monthly Report:** GAS 트리거를 통해 월간 운행기록부(별지 제25호 서식)를 자동 생성.

## 🗂 차량 구분

차량은 `config.json`에서 두 가지 유형으로 관리됩니다.

| 유형                 | 설명                                                          | 설정 키            |
| :------------------- | :------------------------------------------------------------ | :----------------- |
| **고정 사용자 차량** | 특정 직원 전용 차량. 성명 자동 입력.                          | `fixedUser`        |
| **출장용 차량**      | 여러 직원이 공용으로 사용. 성명 선택 + 방문 거래처 필수 입력. | `businessTripCars` |

## 📈 Expected Impact (B&A)

| 구분          | Before (수기 작성)              | After (PPAP 도입)                         |
| :------------ | :------------------------------ | :---------------------------------------- |
| **기록 방식** | 볼펜으로 종이에 작성            | 스마트폰 QR 스캔 후 입력                  |
| **정확도**    | 오기입 및 누락 가능성 높음      | 자동 계산 및 이상 감지 알림               |
| **정산 시간** | 수동 엑셀 타이핑 (수 시간 소요) | 데이터 즉시 추출 및 월간 리포트 자동 생성 |

## ⚙️ Setup

1. `config.json`에서 직원 목록, 고정 사용자 차량, 출장용 차량 정보를 수정합니다.
2. `clasp push`로 GAS에 코드를 배포합니다.
3. GAS 편집기에서 `setupProperties()` 함수를 **수동으로 한 번 실행**하여 Script Properties에 설정값을 저장합니다.
4. 웹앱을 재배포(New Deployment)하면 변경 사항이 적용됩니다.

> ⚠️ `config.json`을 수정한 경우, `clasp push` 후 반드시 `setupProperties()`를 다시 실행해야 합니다.

---

**Developed by Dongmin Lee** _Improving work efficiency through Small but Powerful DX._

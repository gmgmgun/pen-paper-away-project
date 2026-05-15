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

## 🤖 자동 배포 (GitHub Actions)

브랜치별로 자동 배포되며 운영/테스트 환경이 완전히 분리됩니다.

| 브랜치    | 대상       | 동작                                                          |
| :-------- | :--------- | :------------------------------------------------------------ |
| `main`    | 운영 GAS   | `clasp push -f` + 운영 webapp URL 자동 deploy                 |
| `staging` | 테스트 GAS | scriptId 교체 후 `clasp push -f` + 테스트 webapp URL 자동 deploy |
| 수동 실행 | 선택       | Actions 탭 → 환경 선택 (deploy 건너뛰기 옵션 제공)            |

> ⚠️ **main 머지 = 즉시 운영 반영.** PR 리뷰가 사실상 마지막 게이트입니다. deploy 를 건너뛰고 push 만 하고 싶다면 PR 머지 대신 수동 실행 (`skip_deploy=true`) 을 사용하세요.

### 환경 셋업 (최초 1회)

#### 1. 운영 인증

로컬에서 `clasp login` 후 `~/.clasprc.json` 내용을 복사:

- macOS/Linux: `cat ~/.clasprc.json`
- Windows: `type %USERPROFILE%\.clasprc.json`

GitHub repo → **Settings → Secrets and variables → Actions → New repository secret** 에 `CLASPRC_JSON` 으로 등록.

#### 2. 테스트 환경 구축

테스트 환경은 **운영과 완전히 격리된 별도 GAS 프로젝트 + Sheets 사본** 입니다.

1. **Sheets 사본** : Drive 에서 운영 스프레드시트 우클릭 → "사본 만들기" → 이름을 `PPAP 운행일지 (TEST)` 등으로. URL 의 `/d/<ID>/` 에서 **테스트 Sheets ID** 복사.
2. **테스트 GAS 프로젝트 생성** :
   - script.google.com 접속 → "새 프로젝트"
   - 프로젝트 이름을 `PPAP (TEST)` 로 변경
   - 프로젝트 설정 → **스크립트 ID** 복사
3. **테스트 GAS 의 시간대 설정** : 프로젝트 설정 → 시간대 `Asia/Seoul`
4. **테스트 GAS 의 ScriptProperty 사전 세팅** (운영 ID 폴백 방지):
   - 프로젝트 설정 → 스크립트 속성 → 속성 추가
   - 키: `SPREADSHEET_ID`, 값: 위에서 복사한 **테스트 Sheets ID**
5. **첫 webapp 배포** (수동, 1회만):
   - 테스트 GAS 편집기에서 `setupProperties()` 1회 실행
   - "배포 관리" → 새 배포 → **웹앱** 유형 → 배포
   - 받은 **테스트 URL** 메모, **배포 ID** 복사
6. **GitHub Secret 등록**:
   - `GAS_SCRIPT_ID_TEST` : 위 2번에서 복사한 테스트 스크립트 ID
   - `GAS_DEPLOYMENT_ID_TEST` : 위 5번에서 복사한 테스트 배포 ID

이후 `staging` 푸시는 자동으로 push + 같은 URL 에 새 버전 deploy 까지 수행됩니다.

#### 3. 운영용 배포 ID

`GAS_DEPLOYMENT_ID` : 운영 webapp 의 활성 배포 ID. **자동 배포에 필수**.

1. 운영 GAS 편집기 → "배포 관리" → 활성 배포 항목 → **배포 ID** 복사 (`AKfycb...`)
2. Secret 등록: `GAS_DEPLOYMENT_ID` = 위 값

> Secret 이 비어 있으면 CI 가 push 만 하고 deploy 는 경고 후 건너뜁니다 (실패는 아님).

### 일상 워크플로우

```text
feature 작업 → staging 푸시
            ↓ CI 자동
        테스트 GAS push + 테스트 URL 자동 갱신
            ↓ 사용자: 테스트 URL 에서 검증
        OK 확인
            ↓ PR 생성 → main 머지 (리뷰가 최종 게이트)
            ↓ CI 자동
        운영 GAS push + 운영 URL 자동 갱신
```

### Deploy 만 건너뛰고 싶을 때

코드만 편집기에 올리고 URL 은 그대로 두고 싶을 때 (예: 점진적 배포, 운영 시간대 회피):

1. PR 머지 대신 Actions 탭 → "Run workflow"
2. environment=prod, **skip_deploy=true** 선택 → 실행
3. 나중에 deploy 만 원할 때 다시 Run workflow → skip_deploy=false

### 보호 대상 파일

다음 파일은 CI 에 의해 덮어쓰여지지 않도록 `.claspignore` 에 포함되어 있습니다 — 운영/테스트 GAS 편집기에서 각각 직접 관리:

- `config.html` (직원/차량 설정, 민감 정보)
- `config.json`

> ⚠️ 테스트 GAS 에도 `config.html` 을 별도로 두어야 `setupProperties()` 가 동작합니다. 운영의 `config.html` 을 그대로 복사해 두면 충분.

### 트러블슈팅

- **`CLASPRC_JSON secret is not set`** : secret 등록 누락.
- **`GAS_SCRIPT_ID_TEST secret is not set`** : staging 푸시 전 테스트 환경 셋업 누락.
- **`Could not read API credentials`** : clasp 토큰 만료. 로컬에서 `clasp login --no-localhost` 후 secret 재등록.
- **테스트 GAS 가 운영 Sheets 를 건드림** : 위 2번 4단계(`SPREADSHEET_ID` 사전 세팅)를 생략한 경우. ScriptProperty 확인 후 `setupProperties()` 재실행.
- **`clasp push -f` 가 일부 파일을 삭제** : 해당 파일을 `.claspignore` 에 추가.

---

**Developed by Dongmin Lee** _Improving work efficiency through Small but Powerful DX._

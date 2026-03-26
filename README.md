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
회사의 소중한 자산인 법인 차량 2대를 관리함에 있어 기존 방식은 다음과 같은 Pain Point가 있었습니다.
* **불편한 기록:** 매번 차에 비치된 장부를 꺼내 수기로 작성해야 하는 번거로운 프로세스.
* **데이터 불일치:** 계기판 숫자를 잘못 적거나, 바쁜 일정으로 인해 기록을 누락하는 사례 발생.
* **관리 오버헤드:** 정산을 위해 수기 장부 데이터를 다시 엑셀로 옮기는 단순 반복 업무 발생.

## 🛠 Tech Stack & Architecture
가장 빠르고 비용 효율적인 기술 스택을 선택했습니다.
* **Entry:** QR Code (차량 대시보드 부착)
* **Input:** Google Forms (Mobile Optimized)
* **Database:** Google Sheets
* **Logic:** Google Apps Script (GAS)

## ✨ Key Features
* **Scan & Go:** QR 코드 스캔 한 번으로 즉시 기록 페이지 접속.
* **Auto Calculation:** $Distance = Final - Initial$ 수식을 통해 주행 거리 자동 산출.
* **Photo Evidence:** 계기판 사진 업로드 기능을 통해 데이터의 무결성 확보.
* **Cloud Sync:** 모든 데이터는 구글 시트에 실시간 저장되어 별도의 타이핑 작업 불필요.

## 📈 Expected Impact (B&A)
| 구분 | Before (수기 작성) | After (PPAP 도입) |
| :--- | :--- | :--- |
| **기록 방식** | 볼펜으로 종이에 작성 | 스마트폰 QR 스캔 후 입력 |
| **정확도** | 오기입 및 누락 가능성 높음 | 자동 계산 및 사진 증빙 가능 |
| **정산 시간** | 수동 엑셀 타이핑 (수 시간 소요) | 데이터 즉시 추출 (1분 내외) |

---
**Developed by Dongmin Lee** *Improving work efficiency through Small but Powerful DX.*

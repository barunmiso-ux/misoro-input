# misoro-input — 일일 보고 입력 웹앱 (인계 문서 / 컨텍스트 복구용)

## 1. 한 줄 요약
미소로한의원 10개 지점의 일일 보고를 Streamlit 폼으로 받아 구글 시트에 자동 기록.
기존 잔디·Teams 메신저 자유 텍스트 + GAS 파싱 방식을 대체.
배포 URL 예시: `misoro-input.streamlit.app/?branch=분당`

## 2. 왜 만드는가
메신저 + GAS 방식의 한계: 자유 텍스트 파싱 깨짐(한글 "피"→`&#54588;`로 카운트 폭주),
재작성 시 누적(덮어쓰기 안 됨), 댓글 dedupe 추측, 두 메신저 이중 관리.
→ 구조화된 입력 폼으로 전부 해결.

## 3. 결정 사항
| 항목 | 결정 |
|------|------|
| 기술 스택 | Streamlit (Python) + gspread |
| 배포 | Streamlit Cloud (별도 리포 misoro-input) |
| 데이터 백엔드 | 기존 구글 시트 유지 (DB 마이그레이션 안 함) |
| 첫 범위 | 분당 1개 지점 PoC → 검증 후 확장 |
| 메신저 | 폐기 — 전부 웹앱으로 통일 |
| 인증 | URL 파라미터 (`?branch=분당`), 별도 로그인 없음 |
| 직원 환경 | Chrome 앱 모드 (`chrome.exe --app=URL`) |

## 4. 기존 자산
### misoro-dashboard (참고 리포)
- `yongpari86/misoro-dashboard` (main), 배포: https://misoro-dashboard.streamlit.app/
- 인증 패턴: `data_loader._get_gspread_client()` (로컬 service_account.json 우선, 클라우드 st.secrets)
- 서비스 계정: `rank-check-483711@rank-check-483711.iam.gserviceaccount.com`
  - 입력 앱은 **쓰기** 필요 → 일일 보고 시트에 이 이메일을 **편집자**로 추가해야 함
  - (대시보드는 readonly 라 충돌 없음)

### 기존 GAS (Code_v13.gs) — Python 포팅 대상 로직
- 시트 ID: `1uTkikVDCUfVry6l-GX2L8yvlhRSVFdUtEwlubrhpZLk`
- `processSubmission`, `buildColumnMap`, `findDateRow`, `upsertHomeworkRow`, `countItems`, `countDBTags`
- 자유 텍스트 파싱(`parseEntry`)은 폼 입력으로 대체 → 불필요

### 시트 구조
- **일일 시트**: `[YY.MM] 지점별 주간 문의내역` (예: `[26.06] ...`)
  - 지점별 컬럼: 날짜 / 내원 / 워크인 / 네이버 / 카카오 / 전화 / 홈페이지 / DB
  - 폼의 "초진" 소계 → "내원" 컬럼에 기록
- **카페숙제체크 탭**: 날짜 | 지점 | 건강_카페 | 건강_홈페 | 일상게시글 | 댓글수 | 댓글URL | 입력자 | 기록시각
  - 같은 (날짜, 지점) 행 있으면 덮어쓰기 (upsert)
- **휴진관리 탭**: 휴진일 정의 (입력 앱 미사용)
- **지점설정 탭**: 메신저 Webhook (폐기 후 미사용)

### 10개 지점
대전, 부산, 분당, 수원, 안산, 영등포, 인천, 전주, 천안, 평택

## 5. 폼 항목
- **초진·문의** 6개 섹션(초진/네이버/카카오/전화/홈페이지/워크인): 각각 "병명+인원" 동적 리스트.
  초진만 항목별 DB 체크박스. 인원은 number_input(숫자만).
- **활동**: 건강 칼럼(카페 URL, 홈페 URL), 카페 일상 게시글 URL, 카페 댓글 URL 동적 리스트.

### UX 규칙
1. URL 파라미터로 지점 자동 인식 (없으면 선택 박스)
2. 날짜 기본값 오늘, 과거 7일 이내 허용
3. 인원=숫자만, URL=https:// 검증 (폭주 사고 차단)
4. 카페·홈페 URL 선택 입력
5. 같은 날짜+지점 보고 있으면 prefill 수정 모드 (Phase 2)
6. 제출 후 요약 메시지
7. DB는 초진 항목별 체크박스 → 체크된 인원 합이 DB 컬럼

## 6. 현재 구현 상태 (Phase 1 완료)
```
misoro-input/
├── app.py              ← 폼 UI + 검증 + 제출 (USE_MOCK=True, 시트 쓰기 mock)
├── config.py           ← 시트 ID/지점/섹션 정의, daily_sheet_name() 동적 년월
├── sheet_writer.py     ← _get_gspread_client() + write_submission() (Phase 2 TODO)
├── requirements.txt
├── .streamlit/config.toml          ← 하늘색 테마
├── .streamlit/secrets.toml.example ← 서비스 계정 키 양식
├── .gitignore
└── README.md
```
- 로컬 `python -m streamlit run app.py` 부팅 확인됨 (HTTP 200)
- 폼/검증/요약 동작, 시트 쓰기는 mock

## 7. 다음 단계 (Phase 2)
1. `sheet_writer.write_submission()` 본문 구현 (GAS processSubmission 포팅)
   - 탭 찾기 → 컬럼맵 → 날짜행 → 소계/DB setValue (batch) → 카페숙제 upsert
   - **gspread 분당 100회 한도**: 제출 1회당 API 호출 5회 이내
2. `read_existing()` 구현 (수정 모드 prefill)
3. 시트에 서비스 계정 편집자 추가 + 로컬 service_account.json 배치
4. `app.py` 의 `USE_MOCK = False`
5. GitHub 푸시 + Streamlit Cloud 배포 (Secrets 등록)

## 8. 로드맵
Phase 1(완료): 리포 셋업·폼 UI·로컬 테스트 / Phase 2: 시트 쓰기·분당 배포 /
Phase 3: 분당 1~2주 운영·버그수정 / Phase 4: 9개 지점 확장(URL만) /
Phase 5: 메신저 폐기 / Phase 6: 미제출 알림·주간 집계 방식 결정.

## 9. 미결정 (작업 중 결정 가능)
임시 저장(드래프트) 기능 / 수정 권한 제한 / 모바일 최적화 검증 /
DB 마이그레이션 / 미제출 알림 채널(잔디·이메일·Teams).

## 10. 헷갈리는 포인트
- **년/월 하드코딩 금지**: GAS 는 YEAR:2026 박힘. 신규 앱은 `datetime` 동적 (config.daily_sheet_name).
- **인코딩 사고 방지**: 숫자=number_input, URL=text_input+검증. 자유 텍스트 받지 말 것.
- **DB**: 초진 한정. 체크된 항목 인원 합이 DB 컬럼 (기존 `<DB>` 인라인 태그 대체).
- **같은 시트 공유**: dashboard(readonly)와 충돌 없음. 입력 앱만 쓰기.

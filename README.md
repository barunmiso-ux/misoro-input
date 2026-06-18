# misoro-input — 미소로한의원 일일 보고 입력 웹앱

10개 지점이 매일 올리는 일일 보고를 Streamlit 폼으로 받아 구글 시트에 기록한다.
기존 잔디·Teams 자유 텍스트 + GAS 파싱 방식을 대체 (파싱 오류·카운트 폭주 방지).

## 현재 상태 (Phase 2 완료)
- ✅ 폼 UI (초진·문의 6개 섹션 동적 입력 + 활동 섹션)
- ✅ 입력 검증 (인원 = 숫자만, URL = https:// 검증)
- ✅ URL 파라미터 지점 인식 (`?branch=분당`)
- ✅ 제출 요약 메시지
- ✅ **실제 구글 시트 쓰기** (`USE_MOCK=False`) — 분당 E2E 테스트 완료
- ✅ 기존 제출 감지 + 활동 칸 prefill (수정 모드)
- ⏳ Streamlit Cloud 배포 (다음 단계)

### 일일 시트 구조 (실측)
상단 요약(1~15행)은 **수식**(자동 합산) → 직접 쓰지 않음.
실제 입력 대상은 **Daily Report 블록**: 행=날짜(B열), 열=지점별
[내원 워크인 네이버 카카오 전화 홈페이지 (DB)]. 분당=O~U열(분당만 DB 열).
일별 값을 쓰면 상단 요약/월합계가 수식으로 자동 갱신됨.

## 로컬 실행
```bash
pip install -r requirements.txt
streamlit run app.py
```
브라우저에서 `http://localhost:8501/?branch=분당`

## 시트 쓰기 활성화 (Phase 2)
1. `service_account.json` (서비스 계정 키)을 프로젝트 루트에 둔다.
   - 또는 `.streamlit/secrets.toml` 에 `[gcp_service_account]` 등록 (`.example` 참고)
2. 일일 보고 시트에 서비스 계정 이메일을 **편집자**로 추가:
   `rank-check-483711@rank-check-483711.iam.gserviceaccount.com`
3. `sheet_writer.write_submission()` 본문 구현 (GAS `processSubmission` 포팅)
4. `app.py` 의 `USE_MOCK = False`

## 파일 구조
| 파일 | 역할 |
|------|------|
| `app.py` | Streamlit 폼 UI + 검증 + 제출 |
| `config.py` | 시트 ID, 지점, 섹션 정의 (년/월 동적) |
| `sheet_writer.py` | gspread 인증 + 시트 쓰기 (Phase 2) |
| `.streamlit/config.toml` | 테마 |

## 배포 (Streamlit Cloud)
- 리포 연결 후 Secrets 에 `[gcp_service_account]` 등록
- 직원 사용: Chrome 앱 모드 바로가기 `chrome.exe --app=https://.../?branch=분당`

자세한 배경·로드맵은 `START.md` 참고.

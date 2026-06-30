# 바코드 산책(walk) 스캐너 — 버셀 이전 설계

날짜: 2026-06-30

## 배경 / 문제
- 기존: Google Apps Script(GAS) 웹앱이 HTML을 직접 서빙. 폰 카메라로 바코드 스캔.
- 문제: GAS는 HTML을 `*.googleusercontent.com` 교차출처 iframe 안에서 띄움 → 크롬 보안정책(Permissions Policy)상 그 iframe 안에서는 카메라(getUserMedia)가 차단됨.
- 증거: 권한 팝업 주소가 `n-...googleusercontent.com`, 결과 `NotAllowedError: Permission denied`.
- 결론: GAS 주소 안에서는 실시간 카메라 부활 불가. 화면을 GAS 밖(최상위 HTTPS 페이지)으로 빼야 함.

## 목표
- 실시간 카메라 스캔 속도 복원.
- 화면(프론트엔드)은 버셀(Vercel, GitHub 연동) 정적 호스팅으로 이전.
- 구글시트는 저장·조회용 데이터 저장소로만 유지. 기존 시트 로직 재사용.
- 조회/통계(report) 페이지는 제거(미사용).

## 구조
```
[폰 크롬] 버셀 index.html (실시간 Quagga 카메라)
      │  JSONP GET: ?action=함수명&payload=JSON&callback=cb
      ▼
[GAS] doGet API 라우터 → 기존 데이터 함수 실행
      ▼
[구글시트] walk / master / wakeup담당자 (변경 없음)
```

## 데이터 연결 (방법 A: 기존 Code.gs 재활용)
- `doGet`을 API 라우터로 교체. JSONP로 응답(`callback({...})`)하여 CORS 우회.
- 저장도 GET+JSONP로 처리(데이터가 짧음) → POST/CORS 복잡성 회피.
- 노출 액션 5개만:
  - `getManagerList` → 담당자 목록
  - `getDataForSearch` → 학생 마스터(검색/자동완성)
  - `addStudent(payload)` → 저장 후 최근 목록 반환
  - `getRecentWalks` → 최근 목록 + 타임아웃(분)
  - `setEndWalkTime(payload)` → 산책 종료시간 기록
- 페이로드 규약: 복합 인자는 `payload`(JSON 문자열) 1개로 전달, 라우터가 `JSON.parse` 후 분배.

## GAS 변경
- `doGet`: HTML 서빙 → API 라우터로 교체.
- 제거: `loadPartialHTML_`, `loadSearchView` (HTML 조각 서빙용, 더 이상 불필요).
- 보존(삭제 안 함, 미사용 상태로 둠): `querywalkData`, `queryRankData`, `getChartData`,
  `getStudentList`, `searchStudents`, `getStudentStats`, `getAllwalkData`, `parseDateTime`,
  `getWeekStart`, `testGetSheetInfo`, `testGetData`, `jandi`.
- 배포: 새 버전, 실행=나, 액세스=누구나. `/exec` URL 확보.

## 프론트엔드 (버셀 정적)
- `index.html` 하나. 구성:
  - 담당자 선택 모달
  - `#interactive` 라이브 카메라 뷰 + 시작/중지 버튼
  - 검색 폼 + 자동완성(`#searchCode`, `#autocomplete-list`, `#searchInput`, `#tblMessage`)
  - 최근 산책 표(`#searchResults`) + 행 템플릿(`#rowtemplate`) — 기존 search.html 표를 인라인
- 모든 서버호출은 `gasCall(action, payload)` JSONP 헬퍼로 통일.
- 맨 위 `GAS_API_URL` 상수에 본인 GAS `/exec` 주소 설정.
- 라이브 스캔 복원: `getUserMedia(environment)` 권한 → `Quagga.init` LiveStream → `onDetected`(9초 중복방지) → `search()`.
- 제거: 미사용 UI(`studentCard`, `statsDisplay`)와 해당 참조.

## 트레이드오프 / 주의
- JSONP는 주소를 아는 누구나 호출 가능. 단 현재 GAS도 '누구나 접근'이라 보안 수준 동일(신규 악화 없음). 추후 간단한 키로 강화 가능.
- 실시간 카메라·시트 저장은 실제 폰/배포에서 최종 확인 필요(수동 테스트).

## 배포 절차
1. Code.gs를 API 버전으로 교체 → 새 버전 배포(실행=나, 액세스=누구나) → `/exec` 복사
2. `index.html`의 `GAS_API_URL`에 붙여넣기
3. GitHub push → 버셀 자동 배포 → 폰에서 버셀 주소 열기 → "홈 화면에 추가"

# 출석체크 앱

교사용 출석 관리 PWA — 아침자습 + 수업 출결, 구글시트 연동

## 파일 구조
```
attendance/
├── index.html       ← 메인 앱
├── manifest.json    ← PWA 설정
├── sw.js            ← 서비스워커 (오프라인)
└── apps_script.js   ← 구글 Apps Script 코드 (참고용)
```

## 설치 방법

### 1. GitHub Pages 배포
1. `chemistry12` 저장소에 `attendance/` 폴더 업로드
2. `index.html`, `manifest.json`, `sw.js` 3개 파일
3. 접속: `https://chemistry12.github.io/attendance/`

### 2. 구글 Apps Script 연동
1. [sheets.new](https://sheets.new) → 새 스프레드시트 생성
2. URL에서 ID 복사 (`.../spreadsheets/d/[여기]/edit`)
3. 확장 프로그램 → Apps Script
4. `apps_script.js` 내용 붙여넣기
5. `SHEET_ID` 값을 복사한 ID로 교체
6. 배포 → 새 배포 → 웹앱 → 액세스: **모든 사용자** → 배포
7. 배포 URL 복사
8. 앱 하단 URL 입력란에 붙여넣기 → 저장

### 3. 모바일 앱 설치
- **iOS**: Safari → 공유 → 홈 화면에 추가
- **Android**: Chrome → 메뉴 → 앱 설치

## 사용법
- **전체 출석**: 상단 녹색 버튼 한 번으로 전원 출석
- **개별 입력**: 출석/결석/지각 모드 선택 → 학생 카드 탭
- **사유 입력**: 모바일 길게 누르기 / PC 우클릭
- **시트 저장**: 하단 저장 버튼 (날짜별 열 자동 생성)

# DRT · 재고관리

리사이클링팀 재고관리 시스템 (배터리 리사이클링 로트 추적). 두 가지 방식이 나란히 배포돼 있어요.

## 1안 · GitHub Pages + data.json (`/`)

- `site/index.html` — 앱 전체 (React + htm, 외부 CDN 없이 자체 포함)
- `site/data.json` — 팀 공유 데이터. 이 파일만 바뀌면 사이트를 새로 빌드하지 않아도 열려있는 화면에 20초 안에 자동 반영돼요.
- **데이터 갱신 방법**: 사이트에서 설정(⚙️) → PIN으로 잠금 해제 → 재고 입력/수정 → "팀 공유용 내보내기"로 JSON 받기 → 그 파일 내용을 `site/data.json`에 반영해서 커밋·푸시 (또는 Claude에게 전달).

## 2안 · 실시간 연동 (`/v2/`)

- `site/v2/index.html` — Firebase Realtime Database와 브라우저가 직접 통신해요 (REST API + Server-Sent Events, SDK/CDN 없음).
- 데이터베이스: `drt-stock-status` 프로젝트의 Realtime Database, 경로 `/inventory`.
- **데이터 갱신 방법**: 사이트에서 PIN으로 잠금 해제 → 재고 입력/수정 → 저장하는 즉시(1초 이내) 다른 사람이 열어둔 화면에도 자동 반영돼요. 내보내기·전달 절차가 필요 없어요.
- 보기 비밀번호 / 편집 비밀번호를 서로 다르게 설정해서 팀원(보기)과 관리자(편집)를 구분해요.

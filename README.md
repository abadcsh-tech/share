# devotion-share

설교·강의 기반 웹 산출물을 **GitHub Pages로 배포하는 저장소**입니다 (원격 레포명: `share`).
콘텐츠를 생성하는 로직(스킬)은 이 레포에 없고, 여기에는 완성된 산출물만 쌓입니다.

라이브 URL: `https://abadcsh-tech.github.io/share/...`

## 폴더 구조

| 경로 | 내용 |
|---|---|
| `devotional/` | **현행** — 주일 설교 한 편을 청소년부 5일 묵상 인터랙티브 HTML(자기완결 단일 파일)로 변환한 산출물. 파일명은 설교 노트 파일명을 따름 |
| `devotion-10/` | 1세대 5일 묵상 패키지 (슬라이드 HTML + 이미지 + 배포용 PPTX/PDF, 사도신경 강해 10) |
| `youtube/` | 유튜브 촬영용 강의 슬라이드 HTML 시안 |
| `index.html`, `lecture.html` | 레거시 — "MD 파일, 모든 것의 시작" 강의 페이지 (이 레포의 옛 이름 `claude-code-md-explainer`의 유래) |

## 생성 스킬 (이 레포 외부)

- `weekly-devotion` — 설교 .md → 5일 묵상 HTML 생성 후 이 레포의 `devotional/`에 출력·푸시
- `youtube-slides` — 원고 → 촬영용 슬라이드 HTML, `youtube/`에 출력
- 스킬 원본은 별도 저장소 `claude-resources`에서 관리 (`~/.claude/skills/`에 심볼릭 링크)

## 히스토리

- 처음에는 Claude Code × Obsidian 강의 페이지 저장소로 시작 (`claude-code-md-explainer`)
- 2026-07: 공개 정리 — 절대경로가 박힌 빌드 스크립트·Keynote 원본 제거
- 2026-08-13: 로컬 폴더명을 `devotion-share`로 변경 (GitHub 레포명 `share`와 Pages URL은 기존 공유 링크 보호를 위해 유지)

# AI-Elite MIS RPA

빙그레 5개 공장(남양주1·2·김해·광주·논산)의 **MIS 화면 데이터 수집·재가공** 자동화 프로젝트.

BEMS 웹앱(`AI-Elite_Energy-Dashboard-Web`)에서 분리된 독립 프로젝트입니다.

## 역할 분담 (웹앱과의 경계)

```
[이 프로젝트: MIS RPA]                         [웹앱: BEMS]
 MIS 화면 좌표 클릭 → 클립보드 수집             서버 기동 시(main.py):
   ↓                                            auto_sync_once / auto_sync_production_once
 RawDB_*.xlsx (원본 수집)                        ↓  (mtime 변경 시에만)
   ↓  형식 재가공(production_builder / wip_refactoring)
 DB_생산실적.xlsx · DB_재공품.xlsx  ─────►  엑셀 읽어 MySQL 테이블 UPSERT
```

- **접점은 `DB_MIS_DIR` 폴더의 엑셀 파일뿐.** RPA는 DB에 직접 쓰지 않습니다.
- 에너지: RPA가 `RawDB_에너지.xlsx` 생성 → 웹이 startup에 직접 읽어 적재.
- 생산실적: RPA가 `RawDB_생산실적.xlsx` 수집 → `production_builder.build_dataset` 재가공 →
  `DB_생산실적.xlsx` → 웹이 startup에 적재.
- 재공품: RPA가 `RawDB_재공품.xlsx` 수집 → `wip_refactoring` 재가공 → `DB_재공품.xlsx`.

## 구조

```
AI-Elite_MIS_RPA/
├── mis_rpa/
│   ├── config.py               # DB_MIS_DIR 경로 해석 (.env)
│   ├── factories.py            # 공장 코드/도메인 상수
│   ├── production_builder.py   # RawDB_생산실적 → DB_생산실적 재가공 (build_dataset 등)
│   ├── wip_refactoring.py      # RawDB_재공품 → DB_재공품 재가공
│   ├── _common.py              # 클립보드/윈도우/atomic-save 공통 헬퍼
│   ├── production_daily_rpa.py # 생산실적 수집 RPA
│   ├── utility_daily_rpa.py    # 유틸리티(에너지) 수집 RPA
│   ├── wip_daily_rpa.py        # 재공품 수집 RPA
│   ├── build_production_dataset.py  # 생산실적 재가공 CLI
│   ├── run_all_rpa.py          # 3종 RPA 오케스트레이터
│   ├── *_coords.json           # MIS 화면 좌표
│   └── *.bat                   # 실행 래퍼
└── utils/                      # 좌표 측정/클릭 기록 도우미
```

## 설치

```bat
cd /d E:\AI-Elite_MIS_RPA
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
copy .env.template .env
REM .env 의 DB_MIS_DIR 을 웹앱과 동일하게 맞추세요.
```

## 실행

```bat
REM 3종 전체 자동 실행 (생산실적 → 유틸리티 → 재공품)
mis_rpa\전체_RPA_자동실행.bat

REM 개별 실행
python mis_rpa\run_all_rpa.py --date 2026-06-30
python mis_rpa\build_production_dataset.py      # 재가공만
```

> **주의**: RPA 실행 중에는 화면 잠금/화면보호기/모니터 절전/RDP 세션 끊김이 없어야 합니다
> (좌표 클릭 기반). 전원 옵션에서 디스플레이 끄기를 '안 함'으로 설정하세요.

## 웹앱과의 버전 정합

`factories.py`(공장 코드)와 `config.py`(경로 규칙)는 웹앱 `app/domain/factories.py`,
`app/config/paths.py` 에서 복제된 것입니다. 공장 코드 체계가 바뀌면 양쪽을 함께 갱신하세요.
재가공 로직(`production_builder`)은 웹앱의 구 `production_dw_service` build 파이프라인을 이관한 것으로,
웹앱에는 조회 함수(`query_*`)만 남아 있습니다.

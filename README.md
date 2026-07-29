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
- 에너지: RPA가 `RawDB_에너지.xlsx` 수집 → `energy_builder.build_dataset` 재가공 →
  `DB_에너지.xlsx` → 웹이 startup에 적재.
- 생산실적: RPA가 `RawDB_생산실적.xlsx` 수집 → `production_builder.build_dataset` 재가공 →
  `DB_생산실적.xlsx` → 웹이 startup에 적재.
- 재공품: RPA가 `RawDB_재공품.xlsx` 수집 → `wip_refactoring` 재가공 → `DB_재공품.xlsx`.

### 에너지 수집 화면 변경 (2026-07)

주수집 화면이 `유틸리티 일자별 사용량 추이` → `원단위 실적입력(일단위)` 로 바뀌었습니다.
신규 화면은 **단가·비용·COD** 를 함께 제공하지만 **믹스생산량·원단위가 없어** 구 화면을
보완용으로 병행 수집합니다.

```
[원단위 실적입력(일단위)]  냉동전력·공압기·전력량·전력비·전력단가·연료량·연료비·
                           연료단가·용수량·폐수량·원수COD·배출수COD      ─┐
                                                                            ├─► RawDB_에너지.xlsx
[유틸리티 일자별 사용량 추이]  믹스생산량·전력/연료/용수 원단위           ─┘    (행=일자, 열=항목)
                                                                                     │ energy_builder
                                                                                     ▼
                                                                              DB_에너지.xlsx
                                                                            (행=항목, 열=날짜)
```

- 화면 변경 전 `RawDB_에너지.xlsx` 가 맡던 **전치형(행=항목) 역할은 `DB_에너지.xlsx` 로 이관**
  됩니다. 첫 실행 시 `energy_builder.migrate_legacy_rawdb()` 가 자동으로 1회 복사하고
  구 파일은 `backup/RawDB_에너지_legacy_*.xlsx` 로 보관합니다.
- 이관 후 **웹앱 `.env` 의 `ENERGY_SOURCE_XLSX` 를 `DB_에너지.xlsx` 로 지정**해야 합니다.
- 신규 화면은 조회 전용이 아닌 **실적입력** 화면입니다. 좌표가 어긋나면 그리드에 값이
  입력될 수 있으므로, `utility_coords.json` 수정 후에는 반드시 `--dry-run` 으로 검증하세요.

## 구조

```
AI-Elite_AI-Elite-MIS_RPA/
├── AI-Elite-MIS_RPA/
│   ├── config.py               # DB_MIS_DIR 경로 해석 (.env)
│   ├── factories.py            # 공장 코드/도메인 상수
│   ├── production_builder.py   # RawDB_생산실적 → DB_생산실적 재가공 (build_dataset 등)
│   ├── energy_builder.py       # RawDB_에너지 → DB_에너지 재가공 + 에너지 항목 스키마
│   ├── wip_refactoring.py      # RawDB_재공품 → DB_재공품 재가공
│   ├── _common.py              # 클립보드/윈도우/atomic-save 공통 헬퍼
│   ├── production_daily_rpa.py # 생산실적 수집 RPA
│   ├── utility_daily_rpa.py    # 유틸리티(에너지) 수집 RPA — 2개 화면 병행 수집
│   ├── wip_daily_rpa.py        # 재공품 수집 RPA
│   ├── build_production_dataset.py  # 생산실적 재가공 CLI
│   ├── build_energy_dataset.py      # 에너지 재가공 CLI
│   ├── run_all_rpa.py          # 3종 RPA 오케스트레이터
│   ├── *_coords.json           # MIS 화면 좌표
│   └── *.bat                   # 실행 래퍼
└── utils/                      # 좌표 측정/클릭 기록 도우미
```

## 설치

```bat
cd /d E:\AI-Elite_AI-Elite-MIS_RPA
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
copy .env.template .env
REM .env 의 DB_MIS_DIR 을 웹앱과 동일하게 맞추세요.
```

## 실행

```bat
REM 3종 전체 자동 실행 (생산실적 → 유틸리티 → 재공품)
AI-Elite-MIS_RPA\전체_RPA_자동실행.bat

REM 개별 실행
python AI-Elite-MIS_RPA\run_all_rpa.py --date 2026-06-30
python AI-Elite-MIS_RPA\build_production_dataset.py      # 생산실적 재가공만
python AI-Elite-MIS_RPA\build_energy_dataset.py          # 에너지 재가공만

REM 에너지 단독 실행
python AI-Elite-MIS_RPA\utility_daily_rpa.py --ym 2026-07
python AI-Elite-MIS_RPA\utility_daily_rpa.py --dry-run    # MIS 조회만, 엑셀 미기록
python AI-Elite-MIS_RPA\utility_daily_rpa.py --skip-trend # 구 화면(믹스/원단위) 생략
```

> **주의**: RPA 실행 중에는 화면 잠금/화면보호기/모니터 절전/RDP 세션 끊김이 없어야 합니다
> (좌표 클릭 기반). 전원 옵션에서 디스플레이 끄기를 '안 함'으로 설정하세요.

## 웹앱과의 버전 정합

`factories.py`(공장 코드)와 `config.py`(경로 규칙)는 웹앱 `app/domain/factories.py`,
`app/config/paths.py` 에서 복제된 것입니다. 공장 코드 체계가 바뀌면 양쪽을 함께 갱신하세요.
재가공 로직(`production_builder`)은 웹앱의 구 `production_dw_service` build 파이프라인을 이관한 것으로,
웹앱에는 조회 함수(`query_*`)만 남아 있습니다.

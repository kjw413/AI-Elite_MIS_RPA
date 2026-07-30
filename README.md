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
- 에너지: RPA가 `RawDB_에너지.xlsx` 수집 → **재가공 없이** 웹이 startup에 적재.
  (행=일자 tidy 형태라 웹앱 파서가 그대로 읽습니다)
- 생산실적: RPA가 `RawDB_생산실적.xlsx` 수집 → `production_builder.build_dataset` 재가공 →
  `DB_생산실적.xlsx` → 웹이 startup에 적재.
- 재공품: RPA가 `RawDB_재공품.xlsx` 수집 → `wip_refactoring` 재가공 → `DB_재공품.xlsx`.

### 에너지 수집 화면 변경 (2026-07)

수집 화면이 `유틸리티 일자별 사용량 추이` → `원단위 실적입력(일단위)` **한 화면**으로
바뀌었습니다. 신규 화면은 **단가·비용·COD** 를 함께 제공합니다.

```
[원단위 실적입력(일단위)]  냉동전력·공압기·전력량·전력비·전력단가·연료량·연료비·
                           연료단가·용수량·폐수량·원수COD·배출수COD
                                    │
                                    ▼
                            RawDB_에너지.xlsx   행=일자, 열=항목, 시트=공장명
                                    │            → 웹앱이 startup 에 직접 적재
```

**단일 파일 · 재가공 없음.** 행=일자 방향은 `DB_생산실적.xlsx` 의 `daily` 시트,
`DB_재공품.xlsx` 와 통일된 형태이고, 웹앱 파서가 이 형태를 그대로 읽습니다
(A열이 날짜라 전치 감지 분기를 타지 않고 머리글 부분매칭만으로 컬럼이 잡힘).

- 한동안 `DB_에너지.xlsx`(행=항목, 열=날짜)를 중간 산출물로 두고 재가공했으나
  2026-07-30 폐지했습니다. 웹앱 `v5_common.PATH_ENERGY_SOURCE` 가 `RawDB_에너지.xlsx`
  를 가리켜야 합니다.
- 신규 화면은 조회 전용이 아닌 **실적입력** 화면입니다. 좌표가 어긋나면 그리드에 값이
  입력될 수 있으므로, `utility_coords.json` 수정 후에는 반드시 `--dry-run` 으로 검증하세요.
- 공장별 그리드 지문(SHA-1)을 비교해 **직전 공장 그리드를 재복사하면 적재를 거부**합니다.
  좌표가 밀려 다른 공장 데이터가 섞이는 사고를 첫 실행에서 잡습니다.

### 믹스생산량·원단위는 수집하지 않습니다 (2026-07-30)

신규 화면에는 믹스생산량·원단위가 없어 한동안 구 화면을 병행 수집했지만, **BEMS 웹앱이
이미 생산량 분모를 생산실적(`production_daily.actual_qty`)으로 쓰고 원단위를 매 조회마다
재계산**하기 때문에(`production_actual_service.overlay_actual_production` — 화면·메일·분석·
예측 전부 적용) MIS 에서 받아올 필요가 없습니다. 구 화면 수집은 제거했습니다.

- `RawDB_에너지.xlsx` 의 믹스생산량·전력/연료/용수 원단위 열은 **과거 값 보존용**으로
  남기며 신규 날짜에는 비어 있습니다. `write_raw` 는 값이 없는 항목을 건드리지 않습니다.
  (BEMS `validate_columns` 가 필수 컬럼으로 검사하므로 열 자체는 유지해야 합니다)
- 검증: 믹스생산량[kg] 과 생산실적 `actual_qty` 합계는 6개 공장 모두 비율 1.000 으로
  일치합니다. 단 일별로는 약 6% 날짜가 어긋납니다(믹스 제조일과 포장일의 시차).

## 구조

```
AI-Elite_AI-Elite-MIS_RPA/
├── AI-Elite-MIS_RPA/
│   ├── config.py               # DB_MIS_DIR 경로 해석 (.env)
│   ├── factories.py            # 공장 코드/도메인 상수
│   ├── production_builder.py   # RawDB_생산실적 → DB_생산실적 재가공 (build_dataset 등)
│   ├── energy_builder.py       # 에너지 항목 스키마 + RawDB_에너지 적재/읽기
│   ├── wip_refactoring.py      # RawDB_재공품 → DB_재공품 재가공
│   ├── _common.py              # 클립보드/윈도우/atomic-save 공통 헬퍼
│   ├── production_daily_rpa.py # 생산실적 수집 RPA
│   ├── utility_daily_rpa.py    # 에너지 수집 RPA — 원단위 실적입력 화면, 일일/과거 겸용
│   ├── wip_daily_rpa.py        # 재공품 수집 RPA
│   ├── build_production_dataset.py  # 생산실적 재가공 CLI
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

REM 에너지 — 일일 수집 (기본: D-1 기준월)
AI-Elite-MIS_RPA\유틸리티_RPA_실행.bat
python AI-Elite-MIS_RPA\utility_daily_rpa.py --ym 2026-07
python AI-Elite-MIS_RPA\utility_daily_rpa.py --dry-run    # MIS 조회만, 엑셀 미기록
python AI-Elite-MIS_RPA\utility_daily_rpa.py --factories 논산,경산   # 특정 사업장만

REM 에너지 — 과거 데이터 수집 (같은 스크립트, 기간만 지정)
python AI-Elite-MIS_RPA\utility_daily_rpa.py --from 2024-01
python AI-Elite-MIS_RPA\utility_daily_rpa.py --from 2024-01 --to 2026-06
python AI-Elite-MIS_RPA\utility_daily_rpa.py --from 2024-01 --resume   # 중단 후 이어서
```

### 과거 데이터 수집

일일 수집과 **같은 스크립트**입니다. 화면·파싱·적재 경로가 완전히 같고 도는 달 수만
다르므로, `--ym`(1개월) 대신 `--from`/`--to`(기간)를 주면 됩니다. `유틸리티_RPA_실행.bat`
을 인자 없이 실행하면 시작월을 물어보고, 비워 두면 일일 수집으로 동작합니다.

- **월 단위 체크포인트** — 매달 수집 직후 적재하므로 중단돼도 앞선 달은 남습니다.
  `--resume` 로 이미 받은 달을 건너뛰고 이어서 실행하세요.
- **덮어쓰기 주의** — 해당 기간의 수집 항목은 화면 값으로 덮어써집니다(신규 화면이 권위
  소스). legacy 항목(믹스생산량·원단위)은 건드리지 않습니다. 실행 전 자동 백업됩니다.
- 여러 달을 지정하면 예상 소요를 보여주고 확인을 받습니다(`--yes` 로 생략).
  소요는 `조회 횟수 × 약 3초` — 예: 2024-01~2026-07 전 사업장이면 186회 ≈ 9분.
  실행 중에는 마우스/키보드를 사용하지 마세요.
- `--factories 논산,경산` (코드 `F40,F50` 도 가능) 로 일부 사업장만 재수집할 수 있습니다.

> **주의**: RPA 실행 중에는 화면 잠금/화면보호기/모니터 절전/RDP 세션 끊김이 없어야 합니다
> (좌표 클릭 기반). 전원 옵션에서 디스플레이 끄기를 '안 함'으로 설정하세요.

## 웹앱과의 버전 정합

`factories.py`(공장 코드)와 `config.py`(경로 규칙)는 웹앱 `app/domain/factories.py`,
`app/config/paths.py` 에서 복제된 것입니다. 공장 코드 체계가 바뀌면 양쪽을 함께 갱신하세요.
재가공 로직(`production_builder`)은 웹앱의 구 `production_dw_service` build 파이프라인을 이관한 것으로,
웹앱에는 조회 함수(`query_*`)만 남아 있습니다.

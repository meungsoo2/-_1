# 후니드 RPA 자동화 프로젝트

현대그린푸드 식재영업1팀 업무 자동화. H-WIS·더존 ERP iU에서 데이터를 받아 엑셀로 가공·보고서화한다.
사용자는 영업 담당자이며 개발자가 아니다. 설명은 짧고 실무 중심으로, 코드는 바로 실행 가능한 완성본으로 준다.

## 전체 흐름
`RUN_RPA.bat` → `★72_Run_Down_all.py`가 4단계를 순서대로 실행
1. H-WIS 다운로드 (`HWIS_DOWN_CHAJU.py`, Playwright)
2. 더존 ERP iU 다운로드 (`ERP_DOWN_CHAJU.py`, pywinauto)
3. 발주 처리 (`AUTO_BalJU_Today.py`)
4. 단가 갱신 (`AUTO_DanGa_Today.py` → `○32Whoneed_Table(단가).py` 호출)

기타: `EXCEL_CHAJU.py`, `Month_Whoneed_3rd.py`, `ERP_1st_Only.py`, 퍼지 매핑(`auto_mapping_v2.py` 계열)

## 반드시 지킬 제약 (가장 중요)
- **Fasoo DRM**: 엑셀 파일을 pandas/openpyxl로 직접 열거나 저장하지 않는다. 모든 엑셀 작업은 xlwings 또는 pywin32 COM으로, 저장은 엑셀 자체 저장으로 해서 DRM을 유지한다.
- DRM 환경에서 `ClearContents`가 막히면 `None` 값 덮어쓰기로 대체한다.
- **COM 연결**: `win32com.client.dynamic.Dispatch()` 사용. gen_py 캐시를 삭제하는 코드를 넣지 않는다. 엑셀 프로세스 종료 직후 바로 Dispatch하면 경합이 생기므로 대기를 둔다.
- 이미 열린 통합문서는 ROT에서 `display_name.lower().endswith(file_name.lower())`로 찾고, ROT 객체를 Workbook으로 취급한다.
- **H-WIS(Nexacro)**: 요소 ID가 세션마다 바뀌므로 `[id*="조각"]` 부분일치 CSS 셀렉터를 쓴다. 콤보박스는 `fill()` 후 `press("Enter")`.
- **더존 ERP iU(WPF)**: pywinauto UIA 패턴 기반으로 조작한다.
- 다른 드라이브 간 파일 이동은 `os.replace` 대신 `shutil.move`.
- 다중 모니터(노트북 + 모니터 3대) 환경이므로 DPI awareness를 설정한다.

## 코드 규칙
- 모든 스크립트는 `main() -> int`를 노출한다 (PyInstaller 패키징용).
- PyInstaller는 `--onedir`로 빌드한다 (`--onefile`보다 빠름).
- 파워쿼리 경로는 `쿼리경로` 시트(쿼리명 / 현재경로 / 대상경로)를 기준으로 `WorkbookQuery.Formula`의 M코드에 주입한다. 이 구조를 바꾸지 않는다.
- 파워쿼리 새로고침은 mtime 비교(`_refresh_state.json`)로 변경 없으면 건너뛴다.
- 퍼지 매핑 로직을 바꾸면 `MAPPING_LOGIC_VERSION`을 올려 캐시를 재생성한다.
- 수정은 부분 패치보다 **함수/파일 단위 전체 교체본**으로 준다.

## VBA 규칙
- 실행 중 `ScreenUpdating`/`EnableEvents`/`Calculation` 끄고 끝에 복원.
- 메시지박스는 마지막에 한 번만.
- 주석에 이모지 금지 (ANSI 인코딩).
- 대량 행 삭제는 Union으로 묶어 한 번에.

## 환경
- 회사 PC: Windows, Excel 2024 LTSC, 엔트라ID 로그인 (가끔 시스템 응답 없음 → 타임아웃·재시도 고려)
- 집 PC: Microsoft 365 Excel — 회사 PC에서 안 되는 기능(최신 함수 등)을 쓰지 않는다.
- 발주 파일 저장 경로: `C:\후니드_ERP발주\26년\{MMDD}\`

## 작업 방식
- 큰 수정은 먼저 계획을 말하고 확인받은 뒤 진행한다.
- 실행 결과 확인 기준(처리 건수, 생성 파일 등)을 로그로 남긴다.
- 지금 결과가 수작업보다 충분히 나으면 과한 최적화는 하지 않는다.
- 파일 삭제, 원본 덮어쓰기는 반드시 먼저 묻는다.

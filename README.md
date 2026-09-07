# XODR Analysis Guide & CSV Automation

![XODR Analyzer 실행 데모](docs/xodr-analyzer-demo.gif)

OpenDRIVE(`.xodr`) 도로망을 Python·Pandas로 자동 분석하고 Excel, CSV, 그래프와 품질검사 결과를 생성하는 Windows 도구입니다. GUI에서 파일을 선택하거나 XODR을 BAT 파일에 드래그앤드롭하면 분석 결과 폴더가 자동 생성됩니다.

> 전체 원본 영상: [MP4 보기](docs/xodr-analyzer-demo.mp4)

## 다운로드 및 참고 자료

- [XODR Analyzer v2 다운로드](./XODR_Analyzer_v2_Fixed.zip)
- [Pandas 기반 XODR 분석 및 자동화 보고서](./Pandas_XODR_Analysis_Report.html)

## 주요 기능

- XODR XML 구조 자동 파싱
- Road, Lane, Geometry, Junction, Signal, Object 데이터 추출
- 도로별 종합 Summary 생성
- Excel 통합 보고서 저장
- 모든 분석표의 CSV 백업 저장
- 도로 길이, Lane 종류, Geometry 종류, 긴 도로, 복잡한 Junction 그래프 생성
- 짧은 Road, 중복 ID, Geometry 길이 불일치, Lane Link 누락 등 품질검사
- 실행 과정과 오류 로그 저장
- 손상된 Excel 파일이 최종 결과로 남지 않도록 무결성 검사 후 교체

## 설치

1. [XODR_Analyzer_v2_Fixed.zip](./XODR_Analyzer_v2_Fixed.zip)을 다운로드합니다.
2. ZIP을 완전히 압축 해제합니다.
3. 최초 한 번 `install_requirements.bat`을 실행합니다.

필요한 Python 패키지는 다음과 같습니다.

```text
pandas
matplotlib
openpyxl
```

Python이 설치되어 있어야 하며 Windows PATH에서 `py` 또는 `python` 명령을 사용할 수 있어야 합니다.

## 실행 방법

### GUI

`open_xodr_analyzer_gui.bat`을 실행합니다.

1. 분석할 `.xodr` 파일을 선택합니다.
2. 분석 시작 버튼을 누릅니다.
3. 완료되면 결과 폴더가 자동으로 열립니다.

### 드래그앤드롭

분석할 `.xodr` 파일을 `run_xodr_analyzer.bat` 위에 끌어다 놓습니다.

### Python 명령행

```powershell
python xodr_analyzer.py "C:\Maps\Pangyo.xodr"
```

출력 폴더를 직접 지정할 수도 있습니다.

```powershell
python xodr_analyzer.py "C:\Maps\Pangyo.xodr" --output "C:\Maps\Pangyo_Report"
```

## 분석 데이터

| 분석표 | 내용 |
|---|---|
| Roads | Road ID, 이름, 길이, Junction 연결 등 |
| Lanes | Lane Section, Lane ID, 종류, 폭, 주행 방향 등 |
| Geometry | Line, Arc, Spiral, Poly3, ParamPoly3와 좌표·길이 |
| Junctions | Connection과 Lane Link 연결 관계 |
| Signals | 신호 ID, 위치, 방향, 국가, 종류와 값 |
| Objects | 도로 객체의 ID, 종류, 위치와 크기 |
| Summary | Road별 Lane·Geometry·Signal·Object 집계 |
| Statistics | 전체 개수, 길이, 평균값과 품질검사 요약 |

## 자동 품질검사

프로그램은 분석과 함께 다음 항목을 검사합니다.

- 길이가 지나치게 짧은 Road
- 일반 Road와 Junction 내부 Road 구분
- Road 길이와 Geometry 길이 합계의 불일치
- Junction Connection의 Lane Link 누락
- 중복 Road ID
- 같은 Lane Section 내부의 중복 Lane ID

검사 결과는 Excel 시트와 CSV로 함께 저장되어 대규모 XODR의 오류 후보를 빠르게 필터링할 수 있습니다.

## 결과 폴더

입력 파일이 `Pangyo.xodr`이면 같은 위치에 다음 폴더가 생성됩니다.

```text
Pangyo_analysis_result/
├─ Pangyo_XODR_Analysis.xlsx
├─ csv/
│  ├─ Roads.csv
│  ├─ Lanes.csv
│  ├─ Geometry.csv
│  ├─ Junctions.csv
│  ├─ Signals.csv
│  ├─ Objects.csv
│  └─ 기타 요약·품질검사 CSV
├─ charts/
│  ├─ 01_road_length_distribution.png
│  ├─ 02_lane_type_count.png
│  ├─ 03_geometry_type_count.png
│  └─ 기타 분석 그래프
├─ analysis.log
└─ error.log (오류 발생 시)
```

XODR에 특정 항목이 없더라도 가능한 분석 결과와 CSV를 먼저 보존하도록 구성했습니다.

## Excel 저장 안정성

XODR Analyzer v2는 Excel 파일을 임시 위치에 먼저 저장한 뒤 다음 검사를 통과한 경우에만 최종 파일로 교체합니다.

1. XLSX ZIP 내부 구조 검사
2. `openpyxl` 재열기 검사
3. 검사 통과 후 최종 XLSX 교체

Excel 보고서를 열어 둔 채 같은 XODR을 다시 분석하면 저장할 수 없습니다. 기존 Excel 파일을 닫고 다시 실행하세요. Excel 저장이 실패하더라도 CSV 백업과 오류 로그는 유지됩니다.

## 활용 목적

- CARLA·SUMO·RoadRunner용 XODR 구조 검토
- 대규모 도로망 통계 및 품질관리
- Road/Lane/Junction 연결 오류 후보 탐색
- 개발·검수 결과의 Excel 보고
- 여러 지역 XODR을 비교하는 자동화 기반 마련

## 저장소 파일

```text
xodr_analyzer.py               분석 엔진
xodr_analyzer_gui.pyw          GUI
open_xodr_analyzer_gui.bat     GUI 실행기
run_xodr_analyzer.bat          드래그앤드롭/명령행 실행기
install_requirements.bat       Python 패키지 설치기
requirements.txt               의존성 목록
XODR_Analyzer_v2_Fixed.zip     배포용 ZIP
Pandas_XODR_Analysis_Report.html  상세 학습·분석 보고서
docs/xodr-analyzer-demo.gif    README 데모
docs/xodr-analyzer-demo.mp4    원본 영상
```

## 향후 확장

- 여러 XODR 폴더 일괄 분석 및 지역별 비교
- 좌표 범위와 비정상 Geometry 자동 시각화
- RoadRunner·SUMO·CARLA 파이프라인 연계
- 보고서 템플릿과 검수 기준 사용자 설정

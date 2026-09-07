# XODR Analyzer v2

이번 버전은 손상된 Excel 파일이 최종 결과로 남지 않도록 수정했습니다.

## 핵심 개선점

- 임시 XLSX에 먼저 저장
- XLSX ZIP 내부 구조 검사
- openpyxl로 다시 열어서 무결성 검사
- 검사 통과 후에만 최종 XLSX 파일로 교체
- 실패해도 CSV는 먼저 저장
- `analysis.log` 및 `error.log` 생성
- 기존 Excel이 열려 있으면 명확한 오류 표시
- 한글 및 제어문자로 인한 Excel 저장 문제 방지

## 최초 한 번

`install_requirements.bat`를 실행합니다.

## 실행 방법 1: GUI

`open_xodr_analyzer_gui.bat`를 실행합니다.

1. XODR 파일 선택
2. 분석 시작
3. 완료 후 결과 폴더 자동 열림

## 실행 방법 2: 드래그 앤 드롭

`.xodr` 파일을 `run_xodr_analyzer.bat` 위에 끌어다 놓습니다.

## 결과 폴더

XODR 파일 옆에 다음 폴더가 생성됩니다.

`파일명_analysis_result`

## 결과물

- 정상 여부를 재검사한 Excel 보고서
- 모든 시트의 CSV 백업
- 그래프 PNG
- analysis.log
- 오류 발생 시 error.log

## 주의

분석 결과 Excel 파일을 열어둔 채 다시 실행하면 저장할 수 없습니다.
Excel을 닫은 뒤 다시 실행하세요.

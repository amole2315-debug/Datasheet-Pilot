# DataSheet Pilot

MPX2 Word datasheet와 HTRI Excel datasheet를 브라우저에서 일괄 수정하는 도구입니다.

## 실행

```bat
Run DataSheet Pilot.bat
```

브라우저가 열리면 MPX2 또는 HTRI 탭을 선택해서 파일을 업로드하고 task를 적용한 뒤 다운로드합니다.

## 파일 구성

- `DataSheet Pilot.html`: 화면과 MPX2 처리 로직
- `pilot_server.js`: 로컬 서버와 HTRI Excel 변환 로직
- `Run DataSheet Pilot.bat`: 실행 파일
- `templates/BPHE.xlsm`: HTRI BPHE 출력 템플릿
- `templates/GPHE.xlsm`: HTRI GPHE 출력 템플릿

## 필요 프로그램

- Node.js
- Microsoft Excel

HTRI 변환은 Excel COM을 사용하므로 Windows에서 Excel이 설치되어 있어야 합니다.

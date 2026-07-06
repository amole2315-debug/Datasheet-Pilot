# Datasheet Correction

Word 데이터시트(`.docx`)를 폴더 단위로 읽고, 항목을 선택해서 일괄 수정하는 프로그램입니다.

## 실행

개발 환경:

```bat
conda activate MPX2
python datasheet_correction_app.py
```

EXE 빌드:

```bat
build_exe.bat
```

빌드 결과:

```text
dist\DatasheetCorrection.exe
```

## 현재 지원 구조

- 폴더 안의 `.docx` Word 파일
- 하위 폴더까지 `.docx` 검색
- 파일명 기준 `GPHE / BPHE`, `simple / standard / internal` 표시
- Word 표의 항목 라벨 기준 수정
- 미리보기 후 `output` 폴더에 `_corrected.docx` 저장

## 현재 지원 항목

- Basic: Customer, Project Name, Contact, Service, Item No., Datasheet No., Date, Designed by
- Duty Requirements: Media, Heat capacity, Temperature inlet/outlet, Mass flow rate, Volume flowrate, Operating Pressure, Pressure drop
- Unit Data: Design Temperature, Design/Test Pressure, Connection Type, O.H.T.C, Heat transfer area, Number of plates, Surface margin
- Product Properties: Density, Specific heat capacity, Thermal conductivity, Viscosity, Latent heat

## 단위 변환

- kcal/h -> kW
- kcal/(m²·h·K) -> W/(m²·K)
- kcal/(m·h·K) -> W/(m·K)
- kcal/(kg·K) -> kJ/(kg·K)
- kcal/kg -> kJ/kg
- bar_G -> bar_A / kPaG / kPaA
- kPa -> bar / bar_G

원본 파일은 직접 수정하지 않고, 결과는 `output` 폴더에 저장합니다.

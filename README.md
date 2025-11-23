# Reserve Training Admission System (RTAS)
예비군 훈련 입소/결산 위한 전자서명·명부 결산·식비/교통/훈련비 자동화 프로그램입니다.  
- GUI 구성요소: PySide6 + qtpy
- 엑셀 처리: pandas, openpyxl
- 이미지 처리(서명): Pillow
- Windows 연동: pywin32
- 파일/데이터 처리: Python(os, sys, json, io, pathlib, time, re 등)

PySide6(LGPL)은 GUI 프레임워크로서 단순 사용되며,
애플리케이션 전체 라이선스(MIT License)에는 전파되지 않습니다.
본 소프트웨어의 핵심 로직, 기능 구현부, 저장·처리·결산 기능은
모두 MIT License 하에 배포됩니다.
---

##  주요 기능
- 훈련일차(1~5일차)별 명부 자동 불러오기
- 식비·교통비 자동 계산 및 시트 분류 처리
- 전자서명 패드 기능 및 서명 이미지 한셀 저장
- 은행 선택창·계좌번호 입력창 UI 제공
- `pandas`, `openpyxl` 기반 한셀 데이터 가공
- 복수의 시트 자동 병합 및 처리

---

## 기술 스택 (Tech Stack)

| 구성 요소 | 사용 기술 |
|----------|-----------|
| GUI | PySide6 + qtpy |
| 데이터 처리 | pandas, openpyxl |
| 이미지 | Pillow(PIL) |
| Windows 연동 | pywin32 |
| Python 버전 | Python 3.10+ 권장 |

---

## 군 반입 승인 관련 정보
본 소프트웨어는 다음의 라이선스 기준을 충족하여  
군 정보체계 반입 시 **타 시스템 소스코드 공개 의무가 발생하지 않습니다.

자체 라이선스: **MIT License (허용적 라이선스)
- 강제 소스코드 공개 조항 없음  
- 군·공공기관·상업 환경 모두에서 자유로운 사용 가능

외부 라이브러리 구성
본 소프트웨어는 아래와 같이 **LGPL 또는 허용적(MIT/BSD) 라이선스**만 포함합니다.

| 라이브러리 | 라이선스 | 비고 |
|------------|----------|------|
| PySide6 | LGPL 3.0 | 앱 전체 공개 의무 없음 |
| qtpy | MIT | 허용적 |
| pandas | BSD 3-Clause | 허용적 |
| openpyxl | MIT | 허용적 |
| Pillow | 허용적 | 이미지 |
| pywin32 | MIT | Windows 연동 |

소스 코드를 대외 공개하도록 의무화 하는 라이선스가 아님

---

## License
This project is released under the **MIT License**.  
자세한 내용은 `LICENSE` 파일을 참고하십시오.



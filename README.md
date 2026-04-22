# HWP Data Injector

Python 기반의 한글(HWP) 문서 자동화 도구입니다. JSON 포맷의 데이터를 한글 템플릿 파일의 지정된 필드(누름틀)에 매핑하여 자동으로 문서를 생성합니다.

## ⚠️ Legal Disclaimer (법적 고지 및 면책 조항)

1. **Trademarks**: 본 프로젝트는 '한글과컴퓨터(Hancom Inc.)'와 어떠한 제휴, 보증, 후원 관계도 없는 독립적인 개인 프로젝트입니다. 'HWP' 및 '한컴오피스'는 한글과컴퓨터의 등록 상표이며, 본 문서에서의 언급은 오직 상호 운용성(Interoperability) 및 기술적 설명 목적으로만 제한됩니다.
2. **Requirements**: 본 스크립트는 Windows 환경에서 한컴오피스가 제공하는 COM API를 외부에서 호출하여 작동합니다. 따라서 본 도구를 사용하기 위해서는 **사용자 본인의 PC에 정당한 라이선스를 취득한 한컴오피스 소프트웨어가 설치**되어 있어야 합니다.
3. **Limitation of Liability**: 본 소프트웨어는 '있는 그대로(AS-IS)' 제공됩니다. 개발자는 이 스크립트를 사용하여 발생하는 문서의 손상, 데이터 유실, 업무상의 차질 또는 기타 어떠한 직간접적인 법적/경제적 손해에 대해서도 책임을 지지 않습니다. 모든 사용 및 결과에 대한 책임은 전적으로 사용자 본인에게 있습니다.
4. **Copyright**: 본 프로젝트의 소스 코드는 라이선스 하에 배포되지만, 사용자가 주입하는 데이터(JSON) 및 템플릿(HWP) 파일의 저작권은 각 사용자에게 있습니다.

## 🚀 Getting Started

1. Windows 환경 및 한컴오피스 설치 확인
2. Python 환경에 `pywin32` 라이브러리 설치 (`pip install pywin32`)
3. `HWP_TEMPLATE_PATH` 및 `DATA_FILENAME` 변수에 본인의 파일 경로 설정
4. 스크립트 실행

## 📄 License
[MIT License](LICENSE)

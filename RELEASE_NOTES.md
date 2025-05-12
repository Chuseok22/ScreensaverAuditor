# 릴리스 노트: 화면보호기 감사 도구 (ScreensaverAuditor) v2.0

## 주요 변경 사항 (2025년 5월)

### 코드 구조 개선 🏗️

- **코드 리팩토링**: 700줄 이상의 코드를 기능별로 분리하여 MVVM 패턴에 맞게 재구성했습니다.
  - `Models`: 데이터 모델 클래스 (ScreensaverEvent, ScreensaverUsage, CommandLineOptions)
  - `Services`: 주요 비즈니스 로직 서비스 (ScreensaverAuditorService, ExcelExporter)
  - `Utils`: 유틸리티 클래스 (CommandLineParser, ConsoleHelper)

### 새로운 기능 🆕

- **`--audit` 명령어 추가**: 기본 설정(최근 7일)으로 간단히 감사를 실행할 수 있습니다.
  ```
  ScreensaverAuditor.exe --audit
  ```

- **JSON 출력 파싱 지원**: PowerShell 명령의 출력을 JSON 형식으로 받아 처리하는 기능이 추가되었습니다.
  - Microsoft JSON 날짜 형식(`/Date(타임스탬프)/`) 특별 처리 로직 구현
  - 다양한 이벤트 메타데이터 추출 기능 개선

### Excel 출력 개선 📊

- **새로운 필드 추가**: Excel 출력에 다음 필드가 추가되었습니다.
  - 이벤트 유형 (화면보호기 시작/종료)
  - 작업 유형 (TaskDisplayName)
  - 활동 ID (ActivityId)
  - 이벤트 제공자 (ProviderName)
  - 감사 결과 (Keywords)

- **시각화 개선**: 
  - 화면보호기 시작 이벤트는 초록색으로 표시
  - 화면보호기 종료 이벤트는 빨간색으로 표시

### 사용성 개선 🧰

- **예외 처리 강화**: 검색 결과가 없을 때 더 상세한 정보를 제공합니다.
  - 검색 기간 표시
  - 사용한 필터 정보 표시
  - 대안 제안

- **JSON 파싱 오류 처리**: JSON 파싱 오류에 대한 예외 처리가 강화되었습니다.

## 업그레이드 안내

- 기존 버전에서 업그레이드 시 별도의 설정 변경 없이 바로 사용할 수 있습니다.
- 이전에 활성화했던 감사 정책은 그대로 유지됩니다.

## 시스템 요구사항

- 이전 버전과 동일한 .NET 6.0 이상 필요
- Windows 운영체제에서만 동작

## 알려진 이슈

- 일부 오래된 Windows 버전에서는 JSON 출력 형식이 지원되지 않을 수 있습니다.
- 매우 큰 규모의 이벤트 로그를 처리할 때 메모리 사용량이 증가할 수 있습니다.

## 향후 계획

- 통계 분석 기능 강화
- 그래픽 사용자 인터페이스(GUI) 추가
- 보고서 형식 추가 (PDF, HTML)
- 대시보드 기능 추가

---

**개발팀**: Chuseok22
**릴리스 날짜**: 2025년 5월 12일

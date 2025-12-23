# API 안전성 진단 및 개선 제안서

본 문서는 `xll-gen/types` 리포지토리의 API 안전성 및 언어 간 불일치 사항을 진단하고, 이에 대한 개선 이력을 기록합니다.

## 1. 진단 요약 (Executive Summary)

C++ 구현체는 주요 메모리 및 호환성 제약을 준수하고 있으나, **일부 메모리 누수 시나리오**와 **정수 오버플로우** 취약점이 발견되었습니다. 또한, Go 구현체의 `DeepCopy` 로직에서 **DoS(Denial of Service) 위험**이 식별되었습니다. 지원 언어 간에는 일부 특수 타입(`AsyncHandle`, `RefCache`)의 처리 불일치가 존재합니다.

## 2. 언어 및 프로토콜 간 불일치 (Inconsistencies)

### 2.1 [Resolved] 에러 코드 매핑 불일치
*   **현상**: `protocol.fbs`의 `XlError` 열거형은 `Null = 2000`으로 정의되어 있으나, Excel 표준 에러 코드는 `0`부터 시작합니다.
*   **조치**: `src/converters.cpp`에 `ToProtocolError` 헬퍼 함수를 추가하여 변환 로직을 일치시켰습니다.

### 2.2 [Major] 미지원 타입 (Missing Types)
*   **현상**: `RefCache`와 `AsyncHandle` 타입이 `protocol.fbs`에 정의되어 있으나, C++ 변환 로직(`AnyToXLOPER12`, `ConvertScalar`)에서 처리되지 않고 `Nil`로 변환됩니다.
*   **진단**: Excel의 비동기 함수(Async UDF) 지원이나 캐싱 기능 사용 시 동작하지 않을 수 있습니다.
*   **제안**: 해당 타입을 지원하거나, 명시적인 에러(`xlerrNA` 등)를 반환하도록 수정해야 합니다.

### 2.3 [Minor] 문자열 길이 제한 (String Limits)
*   **현상**: Go는 길이 제한이 없으나 Excel은 32,767자 제한이 있습니다.
*   **C++ 안전성**: 입력 변환 시 Truncation 및 길이 제한(200KB)을 통해 오버플로우를 방지하고 있습니다.

## 3. 안전성 분석 (Safety Analysis)

### 3.1 C++ (`src/converters.cpp`, `src/utility.cpp`)
*   **취약점 (Vulnerabilities)**:
    1.  **[High] 메모리 누수 (`RangeToXLOPER12`)**: 예외 발생 시 `ScopeGuard`가 `lpmref` 버퍼를 해제하지 않아 메모리 누수가 발생합니다.
    2.  **[Medium] 정수 오버플로우 (`WideToUtf8`)**: 문자열 길이가 `INT_MAX`를 초과할 경우 `int` 캐스팅으로 인한 오버플로우가 발생하여, 잘못된 메모리 접근이 가능합니다.
    3.  **[Medium] 메모리 누수 (`AnyToXLOPER12`)**: `NumGrid` 변환 시 할당 실패 예외가 발생하면 `XLOPER12` 구조체가 누수됩니다.
    4.  **[Medium] API 안전성 (`ConvertGrid`)**: 공개 API가 예외(`std::bad_alloc`)를 외부로 전파하여 호스트 애플리케이션(Excel)을 크래시시킬 수 있습니다.

### 3.2 Go (`go/protocol`)
*   **취약점 (Vulnerabilities)**:
    1.  **[Critical] OOM / DoS Risk (`DeepCopy`)**: `DeepCopy` 메서드가 입력된 데이터 길이를 검증 없이 신뢰하여 메모리를 할당합니다. 조작된 패킷으로 대량의 메모리 할당을 유도하여 서비스를 거부 상태(DoS)로 만들 수 있습니다.
    2.  **[Low] Validation 미강제**: `Validate()` 함수가 존재하지만 데이터 처리 과정에서 호출이 강제되지 않습니다.

## 4. 개선 이력 및 계획 (Action Plan)

### 완료된 개선 (Completed)
1.  **에러 코드 매핑 수정**: `src/converters.cpp` 수정 완료.
2.  **GridToXLOPER12 안전성 강화**: `ScopeGuard` 개선 및 문자열 Truncation 로직 적용 완료.

### 향후 개선 계획 (Proposed Plan)
1.  **[C++] 잔존 메모리 누수 및 오버플로우 수정**:
    *   `RangeToXLOPER12`, `AnyToXLOPER12`의 `ScopeGuard` 위치 및 로직 수정.
    *   `WideToUtf8`에 입력 길이 제한(INT_MAX) 추가.
    *   `ConvertGrid`에 `try-catch` 블록 추가.
2.  **[Go] DeepCopy 안전성 확보**:
    *   `DeepCopy` 및 `Clone` 시 할당 크기(`DataLength`)가 임계값(예: 100MB)을 초과하는지 검증하는 로직 추가.
3.  **[C++] 미지원 타입 처리**:
    *   `AsyncHandle`, `RefCache` 수신 시 `Nil` 대신 명확한 에러 처리 검토.

# API 안전성 진단 및 개선 제안서

본 문서는 `xll-gen/types` 리포지토리의 API 안전성 및 언어 간 불일치 사항을 진단하고, 이에 대한 개선 이력을 기록합니다.

## 1. 진단 요약 (Executive Summary)

C++ 구현체는 Excel의 메모리 및 호환성 제약을 잘 준수하고 있으나, **에러 코드 매핑 불일치**가 발견되었습니다. 해당 문제는 사용자의 정책 결정(Protocol -> Excel 방향은 2000 유지, Excel -> Protocol 방향은 매핑)에 따라 수정되었습니다.

## 2. 언어 및 프로토콜 간 불일치 (Inconsistencies)

### 2.1 [Resolved] 에러 코드 매핑 불일치
*   **현상**: `protocol.fbs`의 `XlError` 열거형은 `Null = 2000`으로 정의되어 있으나, Excel 표준 에러 코드는 `0`부터 시작합니다.
*   **정책 결정**:
    *   **Protocol -> Excel**: `2000`번대 코드를 그대로 Excel로 전달합니다 (Excel에서 이를 수용하도록 설정됨).
    *   **Excel -> Protocol**: Excel의 `0`번대 에러 코드를 Protocol의 `2000`번대 코드로 매핑합니다.
*   **조치**: `src/converters.cpp`에 `ToProtocolError` 헬퍼 함수를 추가하여 Excel -> Protocol 변환 시 `+2000` 오프셋을 적용했습니다.

### 2.2 [Minor] 미지원 타입 (Missing Types)
*   **현상**: `RefCache`와 `AsyncHandle` 타입이 프로토콜에는 존재하나, C++ `AnyToXLOPER12` 변환 로직에서 `Nil`로 처리됩니다.
*   **제안**: 추후 필요 시 명시적인 에러 반환으로 변경 고려.

### 2.3 [Minor] 문자열 길이 제한 (String Limits)
*   **현상**: Go는 길이 제한이 없으나 C++(Excel)은 32,767자 제한이 있습니다.
*   **C++ 안전성**: 입력 변환 시 200KB로 제한하고, Excel 변환 시 32,767자로 잘라내어(Truncation) 오버플로우를 방지하고 있습니다.

## 3. 안전성 분석 (Safety Analysis)

### 3.1 C++ (`src/converters.cpp`)
*   **우수함 (Strong Safety)**:
    *   **Integer Overflow**: `rows * cols` 및 메모리 할당 크기 검사 수행.
    *   **Memory Safety**: `ScopeGuard`를 통한 RAII 패턴 적용.
    *   **Exception Safety**: `try-catch` 블록으로 크래시 방지.

### 3.2 Go (`go/protocol`)
*   **취약점 (Weakness)**:
    *   **OOM Risk**: `DeepCopy` 시 입력 길이를 신뢰하여 메모리를 할당하므로, 악의적인 패킷에 의한 DoS 가능성이 존재합니다.
    *   **Validation**: `Validate()` 호출이 강제되지 않습니다.

## 4. 개선 이력 및 계획 (Action Plan)

1.  **[완료] 에러 코드 매핑 수정**: `src/converters.cpp` 수정 완료.
2.  **[권장] Go 유효성 검증 강화**: `go/protocol/extensions.go`에 문자열 길이 체크 및 `DeepCopy` 전 검증 로직 추가 권장.

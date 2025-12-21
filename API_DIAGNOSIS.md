# API 안전성 진단 및 개선 제안서

본 문서는 `xll-gen/types` 리포지토리의 API 안전성 및 언어 간 불일치 사항을 진단하고 구체적인 개선 방안을 제시합니다.

## 1. 진단 요약 (Executive Summary)

C++ 구현체는 Excel의 메모리 및 호환성 제약을 잘 준수하고 있으나, **치명적인 에러 코드 불일치(Critical Error Code Mismatch)**가 발견되었습니다. 또한 Go 구현체는 대용량 데이터 처리에 대한 방어 코드가 일부 부족하여 DoS(서비스 거부) 공격에 취약할 수 있습니다.

## 2. 언어 및 프로토콜 간 불일치 (Inconsistencies)

### 2.1 [Critical] 에러 코드 매핑 불일치
*   **현상**: `protocol.fbs`의 `XlError` 열거형은 `Null = 2000`으로 정의되어 있습니다. 그러나 C++ 컨버터(`src/converters.cpp`)는 이 값을 별도의 변환 없이 그대로 `XLOPER12.val.err`에 대입합니다.
*   **문제**: Excel의 표준 에러 코드는 `0 (Null)`, `7 (Div0)`, `15 (Value)` 등입니다. 프로토콜의 `2000`번대 값은 Excel에서 인식하지 못하는 비표준 에러 코드로 처리되거나, 잘못된 에러 메시지를 유발합니다. 반대로 Excel에서 발생한 에러(`0`)가 프로토콜로 변환될 때 `XlError` 열거형 범위 밖의 값이 되어 정의되지 않은 동작(UB)을 유발할 수 있습니다.
*   **영향도**: 높음 (High). 에러 처리가 정상적으로 동작하지 않음.

### 2.2 [Minor] 미지원 타입 (Missing Types)
*   **현상**: `RefCache`와 `AsyncHandle` 타입이 프로토콜과 Go 구현체에는 존재하지만, C++ 변환 로직(`AnyToXLOPER12`)에서는 누락되어 있습니다.
*   **결과**: 해당 타입의 데이터가 C++로 넘어올 경우 기본값인 `Nil`로 조용히 변환되어 데이터 소실이 발생합니다.

### 2.3 [Minor] 문자열 길이 제한 (String Limits)
*   **Go**: 문자열 길이에 대한 명시적 제한이 없습니다.
*   **C++**: Excel의 제한인 32,767자로 잘라내며(Truncation), 입력 변환 시 200KB로 입력을 제한합니다.
*   **위험**: Go에서 긴 문자열을 보낼 경우 경고 없이 뒷부분이 잘려 나갑니다.

## 3. 안전성 분석 (Safety Analysis)

### 3.1 C++ (`src/converters.cpp`)
*   **우수함 (Strong Safety)**:
    *   **Integer Overflow**: `rows * cols` 연산 및 메모리 할당 크기 계산 시 `INT_MAX`, `SIZE_MAX` 오버플로우를 철저히 검사합니다.
    *   **Memory Safety**: `ScopeGuard`를 사용하여 예외 발생 시에도 `XLOPER12` 메모리를 안전하게 해제합니다. `try-catch` 블록으로 외부 예외가 Excel 프로세스를 크래시하지 않도록 방어합니다.
    *   **DoS 방지**: 거대한 문자열 입력 시 전체를 변환하지 않고 200KB까지만 처리하여 메모리 고갈 공격을 방어합니다.

### 3.2 Go (`go/protocol`)
*   **취약점 (Weakness)**:
    *   **OOM Risk**: `DeepCopy` 및 `Clone` 메서드가 입력된 벡터의 길이(`Length`)만큼 메모리를 즉시 할당합니다. 악의적으로 조작된 "길이만 큰" 패킷 수신 시 메모리 부족(OOM)을 유발할 수 있습니다.
    *   **Validation**: `Validate()` 메서드가 존재하지만 `Clone` 등의 내부 동작에서 강제되지 않습니다.

## 4. 개선 방안 (Improvement Plan)

### 4.1 에러 코드 매핑 수정 (즉시 실행 권장)
C++ 컨버터(`src/converters.cpp`)에 에러 코드 변환 함수를 추가해야 합니다.

```cpp
// Protocol(2000+) <-> Excel(0+) 매핑
int XlErrorToExcel(protocol::XlError err) {
    return (int)err - 2000;
}
protocol::XlError ExcelToXlError(int err) {
    return (protocol::XlError)(err + 2000);
}
```

### 4.2 타입 누락 처리
`AnyToXLOPER12`의 `switch` 문에 `AsyncHandle`과 `RefCache` 케이스를 추가합니다. 지원하지 않는 경우 명시적인 에러(`xlerrValue` 또는 `#UNSUPPORTED!`)를 반환하도록 변경하여 디버깅을 돕습니다.

### 4.3 Go 유효성 검증 강화
`extensions.go`의 `Validate()` 함수에 문자열 길이(32,767자) 체크 로직을 추가하고, 데이터 직렬화/역직렬화 시 `Validate()` 호출을 권장하는 주석을 추가합니다.

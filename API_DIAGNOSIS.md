# API 안전성 진단 및 개선 제안서

본 문서는 `xll-gen/types` 리포지토리의 API 안전성 현황과 언어 간 불일치를 진단하고, 구체적인 개선 방안을 제시합니다.

## 1. 진단 요약 (Executive Summary)

현재 API는 기본적인 데이터 변환을 지원하나, **DoS(서비스 거부) 공격 취약점**과 **메모리 누수**, 그리고 **언어 간 기능 불일치**가 존재합니다. 특히 Go 언어의 `DeepCopy` 구현체는 입력 데이터 길이를 검증하지 않아 심각한 DoS 위험이 있으며, C++ 변환 로직에는 예외 발생 시 메모리 누수가 발생하는 경로가 확인되었습니다.

## 2. 상세 진단 (Detailed Diagnosis)

### 2.1 안전성 및 보안 취약점 (Safety & Security)

#### [Critical] Go: DeepCopy DoS 취약점
*   **위치**: `go/protocol/deepcopy.go` (`Grid.DeepCopy`, `NumGrid.DeepCopy`)
*   **현상**: 입력된 `DataLength` 값을 신뢰하여 즉시 메모리를 할당합니다 (`make([]..., l)`).
*   **위험**: 공격자가 조작된 패킷(거대 Length)을 전송할 경우, 서버(또는 Go 프로세스)가 즉시 OOM(Out of Memory)으로 종료될 수 있습니다.
*   **참고**: `extensions.go`에 `Validate()` 함수가 존재하나, `DeepCopy` 과정에서 호출되지 않습니다.

#### [Medium] C++: AnyToXLOPER12 메모리 누수 (BUG-004)
*   **위치**: `src/converters.cpp` (`AnyToXLOPER12`, NumGrid 케이스)
*   **현상**: `new XLOPER12[count]` 할당 시 `std::bad_alloc` 예외가 발생하면, 앞서 할당된 `op` 구조체가 해제되지 않고 누수됩니다. `ScopeGuard`가 할당 이후에 선언되어 있어 예외를 방어하지 못합니다.

#### [Medium] C++: ConvertGrid 예외 전파 (AGENTS.md 위반)
*   **위치**: `src/converters.cpp` (`ConvertGrid`)
*   **현상**: `builder.CreateVector` 등에서 메모리 할당 실패 시 `std::bad_alloc` 예외가 그대로 외부(Excel 호스트)로 전파됩니다. 이는 "Top-level converters must wrap logic in try-catch" 가이드라인을 위반합니다.
*   **위험**: 호스트 프로세스(Excel)의 비정상 종료를 유발할 수 있습니다.

#### [Low] C++: WideToUtf8 정수 오버플로우
*   **위치**: `src/utility.cpp` (`WideToUtf8`)
*   **현상**: 입력 문자열의 길이가 `INT_MAX`를 초과할 경우 `(int)` 캐스팅으로 인해 오버플로우가 발생하며, `WideCharToMultiByte` API에 잘못된 길이 인자가 전달될 수 있습니다.

### 2.2 언어 간 불일치 (Inconsistencies)

#### 미지원 타입 (Missing Types)
*   **항목**: `AsyncHandle`, `RefCache`
*   **Go**: 스키마 및 `DeepCopy`에서 지원.
*   **C++**: `AnyToXLOPER12` 변환 시 해당 타입을 무시하고 `Nil`로 변환 (데이터 소실).

#### 에러 코드 처리 (Error Handling)
*   **C++**: 0-1999 범위의 레거시 에러 코드를 프로토콜 에러(2000+)로 자동 매핑.
*   **Go**: 별도의 매핑 로직 없음 (사용자가 직접 2000+ 코드를 사용해야 함).

## 3. 개선 방안 (Improvement Plan)

### 3.1 [Go] DeepCopy 안전성 강화 (즉시 조치 필요)
*   **조치**: `DeepCopy` 메서드 진입 시 `DataLength`가 `Rows * Cols`와 일치하는지, 그리고 전체 크기가 허용 범위(`math.MaxInt32`) 이내인지 검증하는 로직 추가.
*   **방법**: `extensions.go`의 `Validate()` 로직을 `DeepCopy` 내에 통합하거나 선행 호출.

### 3.2 [C++] 메모리 누수 및 예외 처리 수정
*   **조치 1**: `AnyToXLOPER12`의 `ScopeGuard` 선언 위치를 `op` 할당 직후로 이동하여 `new[]` 예외 발생 시에도 `op`가 해제되도록 수정.
*   **조치 2**: `ConvertGrid` 함수 전체를 `try-catch`로 감싸고, 예외 발생 시 빈 그리드나 에러 객체를 반환하도록 수정.

### 3.3 [C++] 입력 검증 강화
*   **조치**: `WideToUtf8` 함수 도입부에 `wstr.size() > INT_MAX` 검증 추가 (오버플로우 방지).

### 3.4 [공통] 일관성 확보
*   **조치**: C++ `AnyToXLOPER12`에서 `AsyncHandle` 수신 시 `Nil` 대신 `xlerrNA` 또는 `xlerrValue`를 반환하여 데이터 소실을 명시적으로 알림 (추후 비동기 UDF 지원 시 구현).

## 4. 이력 (History)
*   **2023-10-XX**: 초기 진단. 에러 코드 매핑 불일치 해결 완료.

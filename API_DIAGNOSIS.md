# API 안전성 진단 및 개선 제안서

본 문서는 `xll-gen/types` 리포지토리의 API 안전성 현황과 언어 간 불일치를 진단하고, 구체적인 개선 방안을 제시합니다.

## 1. 진단 요약 (Executive Summary)

현재 API는 대부분의 주요 보안 취약점(Go DeepCopy DoS, C++ Integer Overflow)을 해결한 상태이나, **RAII 패턴 위반으로 인한 일부 메모리 누수 가능성**과 **언어 간 기능 불일치(비동기 핸들 미지원)**가 남아있습니다. 기존에 치명적이었던 DoS 취약점들은 코드를 확인한 결과 **해결됨(Resolved)**으로 파악되었습니다.

## 2. 상세 진단 (Detailed Diagnosis)

### 2.1 안전성 및 보안 취약점 (Safety & Security)

#### [Resolved] Go: DeepCopy DoS 취약점 (BUG-018)
*   **상태**: **해결됨 (Secure)**
*   **위치**: `go/protocol/deepcopy.go`
*   **분석**: 현재 구현체는 `make` 호출 전 `DataLength * ElementSize`가 실제 버퍼 크기(`rcv._tab.Bytes`)를 초과하지 않는지 검증하는 로직이 포함되어 있습니다. 따라서 악의적인 길이 조작 공격으로부터 안전합니다.

#### [Resolved] C++: WideToUtf8 정수 오버플로우
*   **상태**: **해결됨 (Secure)**
*   **위치**: `src/utility.cpp`
*   **분석**: 문자열 길이가 `INT_MAX`를 초과하는 경우 명시적으로 `std::runtime_error`를 발생시키는 검증 로직이 존재합니다.

#### [Medium] C++: AnyToXLOPER12 메모리 누수 (BUG-017)
*   **상태**: **취약 (Vulnerable)**
*   **위치**: `src/converters.cpp` (`AnyToXLOPER12`, NumGrid Case)
*   **현상**: `op` 구조체를 할당한 후, `lparray`를 `new`로 할당하는 과정에서 예외가 발생하면 `op`를 해제해줄 `ScopeGuard`가 아직 선언되지 않은 상태입니다. 이로 인해 `op` 객체 누수가 발생합니다.
*   **개선안**: `ScopeGuard`의 선언 위치를 `op` 할당 직후로 이동해야 합니다.

#### [Low] C++: RangeToXLOPER12 메모리 누수 (BUG-014)
*   **상태**: **취약 (Vulnerable)**
*   **위치**: `src/converters.cpp` (`RangeToXLOPER12`)
*   **현상**: `lpmref` 메모리 할당 후 예외 발생 시, `ScopeGuard`가 `ReleaseXLOPER12(op)`를 호출하지만, `ReleaseXLOPER12`는 `op`가 가리키는 동적 메모리(`lpmref`)를 해제하지 않습니다.
*   **개선안**: `ScopeGuard` 내에서 `lpmref`를 명시적으로 해제하거나, 커스텀 삭제 로직을 추가해야 합니다.

### 2.2 언어 간 불일치 (Inconsistencies)

#### 미지원 타입 (Missing Types)
*   **항목**: `AsyncHandle`, `RefCache`
*   **Go**: 스키마 및 `DeepCopy`에서 지원하며 정상 동작함.
*   **C++**: `AnyToXLOPER12` 변환 스위치문(`switch`)에 해당 케이스가 없음. `default`로 빠져 `Nil`로 변환됨.
*   **영향**: Go에서 비동기 핸들 등을 보내도 Excel에서는 `Nil`로 수신됨 (데이터 소실).

#### 에러 코드 처리 (Error Handling)
*   **상태**: **일관성 있음 (Consistent)**
*   **분석**: C++ 라이브러리가 Excel 내부 에러(0~1999)와 Protocol 에러(2000+) 간의 매핑을 `ProtocolErrorToExcel` 함수를 통해 올바르게 처리하고 있음을 확인했습니다.

## 3. 개선 방안 (Improvement Plan)

### 3.1 [C++] 메모리 누수 수정 (BUG-017, BUG-014)
*   **AnyToXLOPER12**: `NumGrid` 처리 블록 진입 시 `ScopeGuard`를 즉시 선언하여 `op` 누수 방지.
*   **RangeToXLOPER12**: `ScopeGuard` 람다 함수 내에 `if (op->val.mref.lpmref) delete[] (char*)op->val.mref.lpmref;` 로직 추가.

### 3.2 [C++] 누락된 타입 지원 추가
*   **AnyToXLOPER12**: `AsyncHandle` 및 `RefCache` 케이스를 추가.
    *   현재 Excel API 제한으로 인해 완벽한 대응은 어려우나, 최소한 `xlerrNA` 또는 `xlerrValue`를 반환하여 `Nil`과 구분하거나, 가능한 경우 문자열 표현으로 변환.
    *   제안: `RefCache`는 Key(String)를 반환하고, `AsyncHandle`은 `#ASYNC!` 와 같은 의미 있는 에러나 문자열을 반환.

## 4. 이력 (History)
*   **2023-10-XX**: 초기 진단.
*   **2025-12-23**: 재진단 완료. Go DoS 취약점 및 C++ 오버플로우 해결 확인. 잔여 메모리 누수 및 타입 불일치 식별.

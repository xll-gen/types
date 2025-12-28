# API 안전성 진단 및 개선 제안서

본 문서는 `xll-gen/types` 리포지토리의 API 안전성 현황과 언어 간 불일치를 진단하고, 구체적인 개선 방안을 제시합니다.

## 1. 진단 요약 (Executive Summary)

현재 API는 주요 보안 취약점(Go DeepCopy DoS, C++ Integer Overflow)과 메모리 누수 문제를 모두 해결한 상태입니다. 이전 진단에서 발견된 **메모리 누수(BUG-014, BUG-017)** 및 **타입 불일치(AsyncHandle, RefCache)** 문제 또한 코드 수정과 검증을 통해 해결되었습니다.

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

#### [Resolved] C++: AnyToXLOPER12 메모리 누수 (BUG-017)
*   **상태**: **해결됨 (Secure)**
*   **위치**: `src/converters.cpp` (`AnyToXLOPER12`, NumGrid Case)
*   **분석**: `NumGrid` 처리 블록에서 `op` 할당 직후 `ScopeGuard`를 선언하여, 이후 `lparray` 할당 중 예외가 발생하더라도 `op`가 안전하게 해제되도록 수정되었습니다.

#### [Resolved] C++: RangeToXLOPER12 메모리 누수 (BUG-014)
*   **상태**: **해결됨 (Secure)**
*   **위치**: `src/converters.cpp` (`RangeToXLOPER12`)
*   **분석**: `ScopeGuard` 내에 `if (op->val.mref.lpmref) delete[] ...` 로직이 추가되어, 예외 발생 시 할당된 `lpmref` 메모리가 명시적으로 해제됩니다.

### 2.2 언어 간 불일치 (Inconsistencies)

#### [Resolved] 미지원 타입 (Missing Types)
*   **항목**: `AsyncHandle`, `RefCache`
*   **상태**: **지원됨 (Handled)**
*   **C++ 구현**:
    *   `RefCache`: `AnyToXLOPER12`에서 `Key` 문자열을 추출하여 반환.
    *   `AsyncHandle`: `AnyToXLOPER12`에서 `"#ASYNC!"` 문자열을 반환하여 핸들 존재를 표시.
*   **분석**: Excel `XLOPER12` 타입의 한계 내에서 적절한 문자열 표현으로 매핑하여 데이터 소실 없이 처리하고 있습니다.

#### 에러 코드 처리 (Error Handling)
*   **상태**: **일관성 있음 (Consistent)**
*   **분석**: C++ 라이브러리가 Excel 내부 에러(0~1999)와 Protocol 에러(2000+) 간의 매핑을 `ProtocolErrorToExcel` 함수를 통해 올바르게 처리하고 있음을 확인했습니다.

## 3. 개선 방안 (Improvement Plan)

### 3.1 [완료] C++ 메모리 누수 수정 (BUG-017, BUG-014)
*   **AnyToXLOPER12**: `ScopeGuard` 선언 위치 수정 완료.
*   **RangeToXLOPER12**: `ScopeGuard` 내 `lpmref` 해제 로직 추가 완료.

### 3.2 [완료] 누락된 타입 지원 추가
*   **AnyToXLOPER12**: `AsyncHandle` 및 `RefCache` 케이스 추가 완료.

## 4. 이력 (History)
*   **2023-10-XX**: 초기 진단.
*   **2025-12-23**: 재진단 완료. Go DoS 취약점 및 C++ 오버플로우 해결 확인.
*   **2025-12-28**: 최종 업데이트. 잔여 메모리 누수(BUG-014, 017) 및 타입 불일치 해결 확인. 문서 현행화.

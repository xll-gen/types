# API 안전성 진단 및 개선 제안서

본 문서는 `xll-gen/types` 리포지토리의 API 안전성 및 언어 간 불일치 사항을 진단하고 개선 방안을 제시합니다.

## 1. 진단 요약 (Diagnosis Summary)

전반적으로 C++ 구현체는 Excel의 메모리 제한과 오류 처리를 방어적으로 수행하고 있으나, Go 구현체는 입력값 검증이 일부 누락되어 있거나 C++(Excel)의 제약 사항과 일치하지 않는 부분이 존재합니다.

### 1.1 안전성 분석 (Safety Analysis)

*   **Go (`go/protocol`)**:
    *   **OOM 취약점**: `DeepCopy` 및 `Clone` 메서드가 입력된 FlatBuffer의 `DataLength`를 기반으로 즉시 슬라이스를 할당(`make`)합니다. 검증되지 않은 거대 입력값이 주어질 경우 메모리 부족(OOM)으로 인한 패닉이 발생할 수 있습니다.
    *   **검증 누락**: `Validate()` 함수가 존재하지만, `Clone` 호출 시에는 강제되지 않습니다.
*   **C++ (`src/converters.cpp`)**:
    *   **방어적 프로그래밍**: `GridToXLOPER12` 등에서 `try-catch` 블록과 `ScopeGuard`를 사용하여 메모리 할당 실패(`std::bad_alloc`)를 안전하게 처리(`xltypeErr` 반환)하고 있습니다.
    *   **문자열 처리**: 입력 문자열을 200KB로 제한하고, 변환 후 32,767자(Excel 제한)로 잘라내어 버퍼 오버플로우 및 DoS를 방지합니다.

### 1.2 언어 간 불일치 (Cross-Language Inconsistencies)

| 항목 | Go (Protcol) | C++ (Excel Adapter) | 불일치 위험 |
| :--- | :--- | :--- | :--- |
| **String Length** | 제한 없음 | 최대 32,767자 (Excel 제한), 입력 200KB 클램핑 | Go에서 긴 문자열 전송 시 C++에서 조용히 데이터가 잘림 (Data Truncation). |
| **Grid Size** | `rows * cols <= MaxInt32` | `rows * cols <= int::max` (32bit) | 일치함 (안전). |
| **Allocation** | `make` (패닉 가능) | `new` (예외 처리됨) | Go 서비스의 안정성 저하 가능성. |
| **DeepCopy** | 재귀적 복사 수행 | (해당 없음, 변환만 수행) | - |

## 2. 개선 방안 (Improvement Plan)

### 2.1 Go 패키지 안전성 강화 (단기)

1.  **String 길이 검증 추가**:
    *   `Scalar` 및 `Grid`의 `Validate()` 메서드에 문자열 길이 체크 로직을 추가합니다.
    *   Excel 호환성을 위해 32,767자를 초과하는 경우 경고 또는 에러를 반환하도록 `extensions.go`를 수정합니다.

2.  **DeepCopy 안전장치 마련**:
    *   `DeepCopy` 내부에서 할당 크기가 비정상적으로 큰 경우(예: 100MB 이상) 에러를 반환하거나 패닉을 방지하는 로직은 FlatBuffers 구조상 어렵습니다.
    *   대안으로, `Clone()` 메서드 진입 시 반드시 `Validate()`를 먼저 수행하도록 강제하거나 문서화합니다.

### 2.2 C++ 변환 로직 명확화

1.  **데이터 손실 경고**:
    *   문자열이 32,767자를 초과하여 잘릴 경우, 로그(`OutputDebugString`)를 남기거나, 엄격 모드(Strict Mode)에서는 `xltypeErr`를 반환하는 옵션을 고려해야 합니다.

### 2.3 스키마 정의 강화

1.  **protocol.fbs 명세화**:
    *   주석을 통해 각 필드의 최대 제약 사항(Max Length, Max Rows 등)을 명시하여 클라이언트 구현 시 참고하도록 합니다.

## 3. 실행 계획 (Action Items)

1.  [Go] `go/protocol/extensions.go`: 문자열 길이(32k) 검증 로직 추가.
2.  [Go] `go/protocol/extensions.go`: `DeepCopy` 전 검증 절차에 대한 가이드 추가.
3.  [Doc] `go/protocol/protocol.fbs`: 제약 사항 주석 추가.

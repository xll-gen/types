BUILD_DIR = build
FLATC = flatc

# Try to find flatc in the build directory if it exists
ifneq ("$(wildcard $(BUILD_DIR)/_deps/flatbuffers-build/flatc)","")
    FLATC = $(BUILD_DIR)/_deps/flatbuffers-build/flatc
endif

.PHONY: all build test clean format generate

all: build

build:
	cmake --preset default
	cmake --build --preset default

test:
	ctest --preset default

clean:
	rm -rf $(BUILD_DIR)

format:
	find src include tests -name '*.cpp' -o -name '*.h' | xargs clang-format -i
	go fmt ./...

generate:
	@if [ ! -x "$(FLATC)" ] && ! command -v $(FLATC) >/dev/null; then \
		echo "Error: flatc not found. Please build the project first or install flatbuffers."; \
		exit 1; \
	fi
	$(FLATC) --go -o go/ go/protocol/protocol.fbs
	$(FLATC) --cpp --scoped-enums -o include/types go/protocol/protocol.fbs

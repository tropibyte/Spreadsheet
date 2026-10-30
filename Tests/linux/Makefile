# Build and run the Grid32 test suites on Linux against the Win32 shim.
#
#   make          build everything
#   make check    build and run every suite
#   make clean
#
# See README.md for what the shim does and does not prove.

CXX      ?= g++
GRID32   := ../../Grid32
SHIM     := win32shim
BUILD    := build

# -Wno-reorder / -Wno-unused-parameter: pre-existing in the Grid32 sources;
# left off so a real new warning stands out.
CXXFLAGS := -std=c++17 -g -O2 -Wall -Wextra \
            -Wno-unknown-pragmas -Wno-reorder -Wno-unused-parameter \
            -I$(SHIM) -I$(GRID32)

GRID_SRCS := $(GRID32)/Grid32Mgr.cpp $(GRID32)/Formulator.cpp $(GRID32)/EditWndProc.cpp
STUBS     := $(SHIM)/win32stubs.cpp

SUITES := $(BUILD)/Grid32Tests $(BUILD)/FormulatorTests $(BUILD)/GridMgrTests \
          $(BUILD)/FormulaRecalcTests $(BUILD)/StreamTests

.PHONY: all check clean
all: $(SUITES)

$(BUILD):
	@mkdir -p $(BUILD)

# Header-only helpers; no Grid32 sources needed.
$(BUILD)/Grid32Tests: ../Grid32Tests.cpp | $(BUILD)
	$(CXX) $(CXXFLAGS) -o $@ $<

# The parser, linked against the stub CGrid32Mgr members in the test file.
$(BUILD)/FormulatorTests: ../FormulatorTests.cpp $(GRID32)/Formulator.cpp | $(BUILD)
	$(CXX) $(CXXFLAGS) -o $@ $^

# The real manager, linked against no-op Win32 stubs.
$(BUILD)/GridMgrTests: GridMgrTests.cpp $(STUBS) $(GRID_SRCS) | $(BUILD)
	$(CXX) $(CXXFLAGS) -o $@ $^

# The formula engine through the real manager: recalculation, number formats
# on formula cells, grid-bounded references.
$(BUILD)/FormulaRecalcTests: FormulaRecalcTests.cpp $(STUBS) $(GRID_SRCS) | $(BUILD)
	$(CXX) $(CXXFLAGS) -o $@ $^

# File import/export (GM_STREAMIN / GM_STREAMOUT / GM_CLEARALL), same linkage.
$(BUILD)/StreamTests: StreamTests.cpp $(STUBS) $(GRID_SRCS) | $(BUILD)
	$(CXX) $(CXXFLAGS) -o $@ $^

check: $(SUITES)
	@rc=0; \
	for t in $(SUITES); do \
	  echo "=== $$t ==="; \
	  $$t || rc=1; \
	  echo; \
	done; \
	exit $$rc

clean:
	rm -rf $(BUILD)

# OnlyOffice Macro CLI - Simple Makefile
# For quick builds without CMake

CXX = g++
CXXFLAGS = -std=c++17 -Wall -Wextra -O2 -I./include -I../third_party/duktape/src
LDFLAGS = -lm -ldl -pthread

# Source files
SRCDIR = src
SOURCES = $(wildcard $(SRCDIR)/*.cpp)
OBJECTS = $(SOURCES:.cpp=.o)
TARGET = onlyoffice-macro-cli

# Duktape source (real OnlyOffice version)
DUKTAPE_SRC = ../third_party/duktape/src/duktape.c
DUKTAPE_OBJ = duktape.o

.PHONY: all clean install test examples

all: $(TARGET)

$(TARGET): $(OBJECTS) $(DUKTAPE_OBJ)
	$(CXX) $(OBJECTS) $(DUKTAPE_OBJ) -o $@ $(LDFLAGS)

%.o: %.cpp
	$(CXX) $(CXXFLAGS) -c $< -o $@

$(DUKTAPE_OBJ): $(DUKTAPE_SRC)
	$(CC) -c $< -o $@

clean:
	rm -f $(OBJECTS) $(DUKTAPE_OBJ) $(TARGET)
	rm -f examples/*.out

install: $(TARGET)
	cp $(TARGET) /usr/local/bin/

test: $(TARGET)
	@echo "Running syntax check tests..."
	./$(TARGET) -c examples/spreadsheet-example.js
	./$(TARGET) -c examples/parameterized-macro.js
	./$(TARGET) -c examples/syntax-test.js
	@echo "Running dry-run tests..."
	./$(TARGET) --dry-run examples/spreadsheet-example.js
	./$(TARGET) --dry-run --verbose examples/parameterized-macro.js

examples: $(TARGET)
	@echo "Running example macros..."
	@echo "=== Basic Spreadsheet Example ==="
	./$(TARGET) --verbose examples/spreadsheet-example.js
	@echo ""
	@echo "=== Parameterized Macro Example ==="
	./$(TARGET) --verbose -p name=TestUser -p color=green -p count=3 examples/parameterized-macro.js
	@echo ""
	@echo "=== Syntax Test ==="
	./$(TARGET) -c --verbose examples/syntax-test.js

help:
	@echo "OnlyOffice Macro CLI Build System"
	@echo ""
	@echo "Targets:"
	@echo "  all      - Build the macro CLI tool"
	@echo "  clean    - Clean build artifacts"
	@echo "  install  - Install to /usr/local/bin"
	@echo "  test     - Run basic tests"
	@echo "  examples - Run example macros"
	@echo "  help     - Show this help"
	@echo ""
	@echo "Note: To use real Duktape instead of stub:"
	@echo "  1. Download Duktape from https://duktape.org/"
	@echo "  2. Replace third_party/duktape.c and duktape.h"
	@echo "  3. Or use existing Duktape from OnlyOffice core"
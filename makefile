CXX = g++
CXXFLAGS = -std=c++20 -Wall -g -MMD -MP
SRC_DIR = src
OBJ_DIR = obj
TARGET = pixelbasher

SRCS = $(wildcard $(SRC_DIR)/*.cpp)
OBJS = $(patsubst $(SRC_DIR)/%.cpp, $(OBJ_DIR)/%.o, $(SRCS))

$(TARGET) : $(OBJS)
	$(CXX) $(CXXFLAGS) $(OBJS) -o $(TARGET)

$(OBJ_DIR)/%.o: $(SRC_DIR)/%.cpp
	mkdir -p $(OBJ_DIR)
	$(CXX) $(CXXFLAGS) -c $< -o $@

-include $(OBJS:.o=.d)

.PHONY: check clean
check: $(TARGET)
	mkdir -p test/output ./converted/import ./converted/import-compare  ./converted/export ./converted/export-compare
	rm -rf converted/import converted/import-compare converted/export converted/export-compare
	echo "Running test with python script"
	python3 diff-pdf-page-statistics.py \
		--base_file=fdo69695-1.doc \
		--max_page=1 \
		--minor_differences=false \
		--resolution=75
	cmp --silent ./converted/import/doc/fdo69695-1.doc_import-1.bmp ./converted/expected/fdo69695-1.doc_import-1.bmp && \
	cmp --silent ./converted/export/doc/fdo69695-1.doc_export-1.bmp ./converted/expected/fdo69695-1.doc_export-1.bmp && \
	echo "Test passed" || echo "Test failed"

clean:
	rm -fr $(OBJ_DIR) \
		$(TARGET) \
		./*.csv \
		converted/export \
		converted/export-compare \
		converted/import \
		converted/import-compare

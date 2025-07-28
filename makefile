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
	@mkdir -p $(OBJ_DIR)
	$(CXX) $(CXXFLAGS) -c $< -o $@

-include $(OBJS:.o=.d)

clean:
	rm -fr $(OBJ_DIR) $(TARGET)

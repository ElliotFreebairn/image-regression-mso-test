#ifndef PIXEL_HPP
#define PIXEL_HPP
#include <iostream>

struct Pixel {
  uint8_t blue {0};
  uint8_t green {0};
  uint8_t red {0};
  uint8_t alpha {0};

  std::string to_string() const {
    return "R:" + std::to_string(static_cast<int>(red)) + " " +
           "G:" + std::to_string(static_cast<int>(green)) + " " +
           "B:" + std::to_string(static_cast<int>(blue)) + " " +
           "A:" + std::to_string(static_cast<int>(alpha)) + "\n";
  }

  bool differs_from(const Pixel& other, bool near_edge, int threshold = 40) {
    int avg_diff = (std::abs(red - other.red) + std::abs(green - other.green) + std::abs(blue - other.blue)) / 3;
    threshold = near_edge ? 250 : threshold;
    return avg_diff > threshold;
  }

  bool is_near_white(int white_threshold = 240) {
    return red > white_threshold; // r == g == b because of greyscale
  }
};
#endif

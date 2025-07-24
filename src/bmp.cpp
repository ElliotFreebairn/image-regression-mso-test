//
//
// Copyright the mso-test contributors
//
// SPDX-License-Identifier: MPL-2.0
//
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at http://mozilla.org/MPL/2.0/.

#include "bmp.hpp"

#include <fstream>
#include <cmath>

BMP::BMP(const char* filename, std::string basename) {
  read(filename);
  background_value = get_average_colour();
  non_background_count = get_non_background_pixel_count(background_value);

  std::vector<bool> edge_mask = sobel_edges(*this);
  this->blurred_edge_mask = blur_edge_mask(*this, edge_mask);
}

BMP::BMP(const BMP& other) {
  file_header = other.file_header;
  info_header = other.info_header;
  colour_header = other.colour_header;
  blurred_edge_mask = other.blurred_edge_mask;
  data = other.data;
  red_count = other.red_count;
  yellow_count = other.yellow_count;
  background_value = other.background_value;
  non_background_count = other.non_background_count;
}

void BMP::read(const char* filename) {
  std::ifstream input {filename, std::ios_base::binary};

  if (!input) {
    throw std::runtime_error(std::string("Can't open BMP file: ") + filename);
  }

  input.read((char*)&file_header, sizeof(file_header)); // read file header data into struct
  if (file_header.file_type != 0x4D42) {
    throw std::runtime_error("Not a BMP file");
  }

  input.read((char*)&info_header, sizeof(info_header)); // read info header data into struct
  if (info_header.bit_count != 32) {
    throw std::runtime_error("Needs to be in RGBA format (32 bits), nothing else");
  }
  if (info_header.height < 0) {
    throw std::runtime_error("The program can treat only BMP images with the origin in the bottom left corner!");
  }

  input.seekg(file_header.offset_data, input.beg);

  info_header.size = sizeof(BMPInfoHeader) + sizeof(BMPColourHeader); // we know that its a 32-bit BMP
  file_header.offset_data = sizeof(BMPFileHeader) + sizeof(BMPInfoHeader) + sizeof(BMPColourHeader);
  file_header.file_size = file_header.offset_data;

  int32_t row_stride = info_header.width * info_header.bit_count / 8;
  uint32_t alligned_stride = (row_stride + 3) & ~3; // This rounds up to the nearest multiple of 4
  size_t padding_size = alligned_stride - row_stride;

  std::vector<uint8_t> padding(padding_size);
  data.resize(row_stride * info_header.height);

  // read the pixel data row by row, and handle the padding if its necessary
  for (int y = 0; y < info_header.height; y++) {
    //uint8_t* row_ptr = data.data() + y * row_stride;
    input.read((char*)(data.data() + y * row_stride), row_stride);

    if (padding_size > 0) {
      input.read((char*)padding.data(), padding_size); // read in the padding so next read is correct
    }
  }
  file_header.file_size += data.size() + padding_size * info_header.height;
}


void BMP::write(const char* filename) {
  std::ofstream output{filename, std::ios_base::binary};
  if (!output) {
    throw std::runtime_error("Cannot open file for writing");
  }

  // write the headers
  output.write((const char*)&file_header, sizeof(BMPFileHeader));
  output.write((const char*)&info_header, sizeof(BMPInfoHeader));
  output.write((const char*)&colour_header, sizeof(BMPColourHeader));

  size_t row_stride = info_header.width * info_header.bit_count / 8;
  size_t alligned_stride = (row_stride + 3) & ~3;
  size_t padding_size = alligned_stride - row_stride;

  std::vector<uint8_t> padding(padding_size, 0);

  // write the pixel data row by row
  for (int y = 0; y < info_header.height; y++) {
    //uint8_t* row_ptr = data.data() + y * row_stride;
    output.write((const char*)(data.data() + y * row_stride), row_stride);
    if (padding_size > 0) {
      output.write((const char*)padding.data(), padding_size);
    }
  }
}

int BMP::get_non_background_pixel_count(int background_value) const {
  int non_background_count = 0;
  size_t pixel_count = get_width() * get_height();
  for (size_t i = 0; i < pixel_count; i++) {
    uint8_t gray_value = data[i * 4]; // Assuming 32-bit BMP, gray value is in the first byte

    if (std::abs(gray_value - background_value) > 20) {
      non_background_count++;
    }
  }
  return non_background_count;
}

int BMP::get_average_colour() const {
  int total_gray = 0;
  int32_t stride = info_header.bit_count / 8;
  size_t pixel_count = get_width() * get_height();
  for (size_t i = 0; i < pixel_count; i++) {
    uint8_t gray = data[i * stride];
    // Calculate average gray value
    total_gray += gray;
  }

  return total_gray / pixel_count;
}

void BMP::set_data(std::vector<uint8_t>& new_data) {
  if (new_data.size() != data.size()) {
    throw std::runtime_error("New data size does not match existing data size");
  }
  data = new_data;
}

std::vector<bool> BMP::blur_edge_mask(const BMP& image, const std::vector<bool>& edge_map, int radius) {
  int32_t width = image.get_width();
  int32_t height = image.get_height();
  std::vector<bool> blurred_mask(width * height, false);

  for (int y = 1; y < height - 1; y++) {
      for (int x = 1; x < width - 1; x++) {
          int index = y * width + x;

          // If the current pixel is not an edge, skip it
          if (!edge_map[index])
              continue;

          blur_pixels(x, y, radius, width, height, blurred_mask);
      }
  }
  return blurred_mask;
}

std::vector<bool> BMP::sobel_edges(const BMP& image, int threshold) {
  int32_t width = image.get_width();
  int32_t height = image.get_height();
  const std::vector<uint8_t>& data = image.get_data();
  int32_t pixel_stride = 4;

  // Initalise the result vector to false, indicating no edges found initially
  std::vector<bool> result(width * height, false);

  // Loop through each pixel in the image, skipping the edges to avoid out-of-bounds access
  for (int y = 1; y < height - 1; y++) {
    for (int x = 1; x < width - 1; x++) {

      std::array<int, 2> gradients = get_sobel_gradients(y, x, data, width, pixel_stride);
      int g_x = gradients[0];
      int g_y = gradients[1];
      // Calculate gradient magnitude (clamped to 255)
      int magnitude = std::min(255, static_cast<int>(std::sqrt(g_x * g_x + g_y * g_y)));

      result[y * width + x] = (magnitude >= threshold);
    }
  }
  return result;
}

void BMP::blur_pixels(int x, int y, int radius, int width, int height, std::vector<bool>& mask) {
  for (int dy = -radius; dy <= radius; dy++) {
    for (int dx = -radius; dx <= radius; dx++) {
        int new_x = x + dx;
        int new_y = y + dy;

        // Check if the new coordinates are within bounds
        if (new_x >= 0 && new_x < width && new_y >= 0 && new_y < height) {
            mask[new_y * width + new_x] = true;
        }
    }
  }
}

std::array<int, 2> BMP::get_sobel_gradients(int y, int x, const std::vector<uint8_t>& data, int width, int pixel_stride) {
  // calculate the corners maybe?

  auto index = [&](int row, int col) {
    return (row * width + col) * pixel_stride;
  };

  int top_left = data[index(y - 1, x - 1)];
  int top_mid = data[index(y - 1, x)];
  int top_right = data[index(y - 1, x + 1)];
  int middle_left = data[index(y, x - 1)];
  int middle_right = data[index(y, x + 1)];
  int bottom_left = data[index(y + 1, x - 1)];
  int bottom_mid = data[index(y + 1, x)];
  int bottom_right = data[index(y + 1, x + 1)];

  int g_x = (-1 * top_left) + (-2 * middle_left) + (-1 * bottom_left) +
            (top_right) + (2 * middle_right) + (bottom_right);
  int g_y = (top_left) + (2 * top_mid) + (top_right) +
            (-1 * bottom_left) + (-2 * bottom_mid) + (-1 * bottom_right);

  return {g_x, g_y};
}
//
//
// Copyright the mso-test contributors
//
// SPDX-License-Identifier: MPL-2.0
//
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at http://mozilla.org/MPL/2.0/.

#include <cmath>
#include <algorithm>

#include "pixelbasher.hpp"
#include "pixel.hpp"

BMP PixelBasher::compare_to_bmp(BMP& original, BMP& target, bool enable_minor_differences) {
  int min_width = std::min(original.get_width(), target.get_width());
  int min_height = std::min(original.get_height(), target.get_height());
  int pixel_stride = 4;

  BMP diff(original);

  std::vector<bool> intersection_mask = get_intersection_mask(original, target, min_width, min_height);
  // The diff data is based on the base image, so we create a new data vector to hold the modified pixels
  auto& original_data = original.get_data();
  auto& target_data = target.get_data();
  std::vector<uint8_t> diff_data = original.get_data();

  int original_width = original.get_width();
  int target_width = target.get_width();
  // Loops through the min width and height, if size of images differ slightly.
  for (int y = 0; y < min_height; y++) {
    for (int x = 0; x < min_width; x++) {
      // The intersection mask is based on height and width not stride, so we calculate it separately
      int original_index = (y * original_width + x) * pixel_stride;
      int target_index = (y * target_width + x) * pixel_stride;
      int mask_index = y * min_width + x;

      Pixel original_pixel = Pixel::get_pixel(original_data, original_index);
      Pixel target_pixel = Pixel::get_pixel(target_data, target_index);
      bool near_edge = intersection_mask[mask_index];

      std::vector<uint8_t> bgra = compare_pixels(original_pixel, target_pixel, diff, near_edge, enable_minor_differences);

      for (int i = 0; i < pixel_stride; i++) {
        diff_data[original_index + i] = bgra[i];
      }
    }
  }
  diff.set_data(diff_data);
  return diff;
}

BMP PixelBasher::compare_regressions(BMP& original, BMP& current, BMP& previous) {
  int32_t min_width = std::min(original.get_width(), current.get_width());
  int32_t min_height = std::min(original.get_height(), current.get_height());
  int32_t pixel_stride = 4;

  BMP diff(original);

  // The diff data is based on the base image, so we create a new data vector to hold the modified pixels
  std::vector<uint8_t> diff_data = original.get_data();

  int32_t original_width = original.get_width();
  int32_t current_width = current.get_width();
  int32_t previous_width = previous.get_width();
  for (int y = 0; y < min_height; y++) {
    for (int x = 0; x < min_width; x++) {
      int original_index = (y * original_width + x) * pixel_stride;
      int current_index = (y * current_width + x) * pixel_stride;
      int previous_index = (y * previous_width + x) * pixel_stride;

      Pixel original_pixel = Pixel::get_pixel(original.get_data(), original_index);
      Pixel current_pixel = Pixel::get_pixel(current.get_data(), current_index);
      Pixel previous_pixel = Pixel::get_pixel(previous.get_data(), previous_index);

      std::vector<uint8_t> bgra = compare_versions(current_pixel, previous_pixel, original_pixel);

      for (int i = 0; i < pixel_stride; i++) {
        diff_data[current_index + i] = bgra[i];
      }
    }
  }
  diff.set_data(diff_data);
  return diff;
}


std::vector<bool> PixelBasher::get_intersection_mask(BMP& original, BMP& target, int width, int height) {

  std::vector<bool> original_edge_map = original.get_blurred_edge_mask();
  std::vector<bool> target_edge_map = target.get_blurred_edge_mask();

  std::vector<bool> intersection_mask(width * height, false);
  for (int i = 0; i < width * height; i++) {
    intersection_mask[i] = original_edge_map[i] && target_edge_map[i];
  }
  return intersection_mask;
}

std::vector<uint8_t> PixelBasher::compare_versions(Pixel& current_pixel, Pixel& previous_pixel, Pixel& base) {
  if (current_pixel.is_red() && previous_pixel.is_red()) {
    return colour_pixel(Colour::BLUE); // error has remained
  }

  if (current_pixel.is_red() && !previous_pixel.is_red()) {
    return colour_pixel(Colour::RED); // a regression
  }

  if (!current_pixel.is_red() && previous_pixel.is_red()) {
    return colour_pixel(Colour::GREEN); // a fix
  }
  return base.get_vector();
}


std::vector<uint8_t> PixelBasher::compare_pixels(Pixel& base, Pixel& target, BMP& diff, bool near_edge, bool minor_differences) {
  bool colour_differs = base.differs_from(target, near_edge);

  if (!colour_differs || (near_edge && !minor_differences)) return base.get_vector();

  if (near_edge && minor_differences) {
    diff.increment_yellow_count(1);
    return colour_pixel(Colour::YELLOW);
  }

  diff.increment_red_count(1);
  return colour_pixel(Colour::RED);
}

std::vector<uint8_t> PixelBasher::colour_pixel(Colour colour) {
  switch (colour)
  {
  case Colour::YELLOW:
    return {0, 197, 255, 255};
    break;
  case Colour::RED:
    return {0, 0, 255, 255};
    break;
  case Colour::BLUE:
    return {255, 0, 0, 255};
    break;
  case Colour::GREEN:
    return {0, 255, 0, 255};
    break;
  default:
    throw std::runtime_error("Invalid colour enum passed");
  }
}
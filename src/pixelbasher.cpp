//
//
// Copyright the mso-test contributors
//
// SPDX-License-Identifier: MPL-2.0
//
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at http://mozilla.org/MPL/2.0/.

#include "pixelbasher.hpp"
#include <cmath>
#include <algorithm>

#include "pixel.hpp"

// Compares two BMP images and generates a diff image based on the differences
BMP PixelBasher::compare_to_bmp(BMP& original, BMP& target, bool enable_minor_differences) {
  int32_t min_width = std::min(original.get_width(), target.get_width());
  int32_t min_height = std::min(original.get_height(), target.get_height());
  int32_t pixel_stride = 4;

  BMP diff(original);

  // Retrieve the edge maps for both images using the Sobel algorithm (which pixels are edges)
  std::vector<bool> original_edge_map = sobel_edges(original);
  std::vector<bool> target_edge_map = sobel_edges(target);

  // Blur the edge maps to create a mask that includes nearby pixels (originald on a given radius)
  std::vector<bool> original_blurred_mask = blur_edge_mask(original, original_edge_map);
  std::vector<bool> target_blurred_mask = blur_edge_mask(target, target_edge_map);

  // Create an intersection mask that combines both blurred edge masks, to be used later for comparison
  // This mask will be true for pixels that are edges in both images
  std::vector<bool> intersection_mask(min_width * min_height, false);
  for (int i = 0; i < min_width * min_height; i++) {
    intersection_mask[i] = original_blurred_mask[i] && target_blurred_mask[i];
  }

  // The diff data is based on the base image, so we create a new data vector to hold the modified pixels
  std::vector<uint8_t> diff_data = original.get_data();

  int32_t original_stride = original.get_width();
  int32_t target_stride = target.get_width();
  // Loops through the min width and height, if size of images differ slightly.
  for (int y = 0; y < min_height; y++) {
    for (int x = 0; x < min_width; x++) {
      // Calculate the current index in the data vector for both images
      // The intersection mask is based on height and width not stride, so we calculate it separately
      int original_index = (y * original_stride + x) * pixel_stride;
      int target_index = (y * target_stride + x) * pixel_stride;
      int mask_index = y * min_width + x;

      Pixel original_pixel = Pixel::get_pixel(original.get_data(), original_index);
      Pixel target_pixel = Pixel::get_pixel(target.get_data(), target_index);
      std::array<uint8_t, 4> bgra = {original_pixel.blue, original_pixel.green, original_pixel.red, original_pixel.alpha};

      // If the pixels differ, we need to determine if they are near an edge and apply the appropriate colour
      bool near_edge = intersection_mask[mask_index];
      if (original_pixel.differs_from(target_pixel, near_edge)) {
        if (near_edge) {
          // If minor differences are enabled, we use yellow for pixels near edges
          if (enable_minor_differences)
              bgra = colour_pixel(Colour::YELLOW);
          diff.increment_yellow_count(1);
        } else {
          // If the pixel is not near an edge, we use red for significant differences
          bgra = colour_pixel(Colour::RED);
          diff.increment_red_count(1);
        }
      }
      for (int i = 0; i < pixel_stride; i++) {
        diff_data[original_index + i] = bgra[i];
      }
    }
  }
  // Set the new data to the base image, which now contains the diff
  diff.set_data(diff_data);
  return diff; // Return the modified base image with the diff applied
}

BMP PixelBasher::compare_regressions(BMP& original, BMP& current, BMP& previous) {
  int32_t min_width = std::min(original.get_width(), current.get_width());
  int32_t min_height = std::min(original.get_height(), current.get_height());
  int32_t pixel_stride = 4;

  BMP diff(original);

  // The diff data is based on the base image, so we create a new data vector to hold the modified pixels
  std::vector<uint8_t> diff_data = original.get_data();

  int32_t original_stride = original.get_width();
  int32_t current_stride = current.get_width();
  int32_t previous_stride = previous.get_width();

  for (int y = 0; y < min_height; y++) {
    for (int x = 0; x < min_width; x++) {
      int original_index = (y * original_stride + x) * pixel_stride;
      int current_index = (y * current_stride + x) * pixel_stride;
      int previous_index = (y * previous_stride + x) * pixel_stride;

      Pixel original_pixel = Pixel::get_pixel(original.get_data(), original_index);
      Pixel current_pixel = Pixel::get_pixel(current.get_data(), current_index);
      Pixel previous_pixel = Pixel::get_pixel(previous.get_data(), previous_index);
      std::array<uint8_t, 4> bgra = {original_pixel.blue, original_pixel.green, original_pixel.red, original_pixel.alpha};

      // three cases: PrevLO correct LO wrong = red, PrevLO wrong LO correct = green,  Both wrong = blue, Both correct = none
      if (!previous_pixel.is_red() && current_pixel.is_red()) {
        bgra = colour_pixel(Colour::RED);
      } else if (previous_pixel.is_red() && !current_pixel.is_red()) {
        bgra = colour_pixel(Colour::GREEN);
      } else if (previous_pixel.is_red() && current_pixel.is_red()) {
        bgra = colour_pixel(Colour::BLUE);
      }

      for (int i = 0; i < pixel_stride; i++) {
        diff_data[current_index + i] = bgra[i];
      }
    }
  }
  diff.set_data(diff_data);
  return diff;
}

std::vector<bool> PixelBasher::sobel_edges(BMP& image, int threshold) {
  int32_t width = image.get_width();
  int32_t height = image.get_height();
  const std::vector<uint8_t>& data = image.get_data();
  int32_t pixel_stride = 4;

  // Initalise the result vector to false, indicating no edges found initially
  std::vector<bool> result(width * height, false);

  // Loop through each pixel in the image, skipping the edges to avoid out-of-bounds access
  for (int y = 1; y < height - 1; y++) {
    for (int x = 1; x < width - 1; x++) {

      // Sobel Gx
      int g_x = -1 * data[((y - 1) * width + (x - 1)) * pixel_stride]
               - 2 * data[(y * width + (x - 1)) * pixel_stride]
               - 1 * data[((y + 1) * width + (x - 1)) * pixel_stride]
               + 1 * data[((y - 1) * width + (x + 1)) * pixel_stride]
               + 2 * data[(y * width + (x + 1)) * pixel_stride]
               + 1 * data[((y + 1) * width + (x + 1)) * pixel_stride];

      // Sobel Gy
      int g_y= -1 * data[((y - 1) * width + (x - 1)) * pixel_stride]
               - 2 * data[((y - 1) * width + x) * pixel_stride]
               - 1 * data[((y - 1) * width + (x + 1)) * pixel_stride]
               + 1 * data[((y + 1) * width + (x - 1)) * pixel_stride]
               + 2 * data[((y + 1) * width + x) * pixel_stride]
               + 1 * data[((y + 1) * width + (x + 1)) * pixel_stride];

      // Calculate gradient magnitude (clamped to 255)
      int magnitude = std::min(255, static_cast<int>(std::sqrt(g_x * g_x + g_y * g_y)));

      result[y * width + x] = (magnitude >= threshold);
    }
  }
  return result;
}

std::vector<bool> PixelBasher::blur_edge_mask(const BMP& image, const std::vector<bool>& edge_map) {
  int32_t width = image.get_width();
  int32_t height = image.get_height();
  std::vector<bool> blurred_mask(width * height, false);

  for (int y = 1; y < height - 1; y++) {
      for (int x = 1; x < width - 1; x++) {
          int index = y * width + x;

          // If the current pixel is not an edge, skip it
          if (!edge_map[index])
              continue;

          // If the current pixel is an edge, we apply a blur effect by marking nearby pixels
          int radius = 2;
          for (int dy = -radius; dy <= radius; dy++) {
              for (int dx = -radius; dx <= radius; dx++) {
                  int new_x = x + dx;
                  int new_y = y + dy;

                  // Check if the new coordinates are within bounds
                  if (new_x >= 0 && new_x < width && new_y >= 0 && new_y < height) {
                      blurred_mask[new_y * width + new_x] = true;
                  }
              }
          }
      }
  }
  return blurred_mask;
}

std::array<uint8_t, 4> PixelBasher::colour_pixel(Colour colour) {
  if (colour == Colour::YELLOW) {
    return {0, 197, 255, 255};
  } else if (colour == Colour::RED) {
    return {0, 0, 255, 255};
  } else if (colour == Colour::BLUE) {
    return {255, 0, 0, 255};
  } else if (colour == Colour::GREEN) {
    return {0, 255, 0, 255};
  }
  return {0, 0, 0, 0};
}

//
//
// Copyright the mso-test contributors
//
// SPDX-License-Identifier: MPL-2.0
//
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at http://mozilla.org/MPL/2.0/.

#ifndef PIXELBASHER_HPP
#define PIXELBASHER_HPP

#include <vector>
#include <array>

#include "bmp.hpp"
#include "pixel.hpp"

// PixelBasher class to handle the comparison of BMP images and generate diff images
class PixelBasher
{
public:
  enum Colour
  {
    RED,
    YELLOW,
    BLUE,
    GREEN
  };
  // Compares two BMP images and generates a diff image based on the differences (the diff is applied to the base image)
  BMP compare_to_bmp(BMP &original, BMP &target, bool enable_minor_differences);
  BMP compare_regressions(BMP &original, BMP &current, BMP &previous);

private:
  std::vector<bool> get_intersection_mask(BMP &original, BMP &target, int min_width, int min_height);

  std::vector<uint8_t> compare_versions(Pixel& current_pixel, Pixel& previous_pixel, Pixel& base);
  std::vector<uint8_t> compare_pixels(Pixel& base, Pixel& target, BMP &diff, bool near_edge, bool minor_differences);
  std::vector<uint8_t> colour_pixel(Colour colour);
};
#endif

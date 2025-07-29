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

#include <array>
#include <cstdint>
#include <vector>

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
		DARK_YELLOW,
		BLUE,
		GREEN
	};
	// Compares two BMP images and generates a diff image based on the differences (the diff is applied to the base image)
	static BMP compare_bmps(const BMP &original, const BMP &target, bool enable_minor_differences);
	static BMP compare_regressions(const BMP &original, const BMP &current, BMP &previous);

private:
	static std::vector<bool> get_intersection_mask(const BMP &original, const BMP &target, int min_width, int min_height);
	static std::array<std::uint8_t, 4> compare_pixel_regression(std::array<std::uint8_t, 4> original, std::array<std::uint8_t, 4> current, std::array<std::uint8_t, 4> previous);
	static std::array<std::uint8_t, 4> compare_pixels(std::array<std::uint8_t, 4> original, std::array<std::uint8_t, 4> target, BMP &diff, bool near_edge, bool minor_differences);
	static std::array<std::uint8_t, 4> colour_pixel(Colour colour);
};
#endif

//
//
// Copyright the mso-test contributors
//
// SPDX-License-Identifier: MPL-2.0
//
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at http://mozilla.org/MPL/2.0/.

#ifndef PIXEL_HPP
#define PIXEL_HPP

#include <array>
#include <cstdint>
#include <iostream>
#include <vector>


#include "bmp.hpp"

struct Pixel
{
	std::uint8_t blue{0};
	std::uint8_t green{0};
	std::uint8_t red{0};
	std::uint8_t alpha{0};

	static std::array<std::uint8_t, pixel_stride> get_vector(const std::uint8_t* src_row)
	{
		return {src_row[0], src_row[1], src_row[2], src_row[3]};
	}

	// Compare this pixel with another pixel and return true if they differ
	static bool differs_from(std::array<std::uint8_t, 4> original, std::array<std::uint8_t, 4> target, bool near_edge, int threshold = 40)
	{
		int avg_diff = std::abs(original[2] - target[2]);
		threshold = near_edge ? 180 : threshold;
		return avg_diff > threshold;
	}

	static bool is_red(std::array<uint8_t, pixel_stride> pixel)
	{
		return pixel[0] == 0 && pixel[1] == 0 && pixel[2] == 255;
	}
};
#endif

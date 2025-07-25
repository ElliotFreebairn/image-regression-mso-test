//
//
// Copyright the mso-test contributors
//
// SPDX-License-Identifier: MPL-2.0
//
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at http://mozilla.org/MPL/2.0/.

#ifndef BMP_HPP
#define BMP_HPP

#include <array>
#include <cstddef>
#include <cstdint>
#include <iostream>
#include <vector>

#pragma pack(push, 1)
struct BMPFileHeader
{
	uint16_t file_type; // BM
	uint32_t file_size;
	uint16_t placeholder_1;
	uint16_t placeholder_2;
	uint32_t offset_data; // this is the start position of pixel data (bytes)
};

struct BMPInfoHeader
{
	uint32_t size;  // Size of the header (bytes)
	int32_t width;  // in pixels
	int32_t height; // in pixels
	uint16_t planes;
	uint16_t bit_count; // useful to check if file is RGBA or RGB
	uint32_t compression;
	uint32_t size_image;
	int32_t x_per_meter;
	int32_t y_per_meter;
	uint32_t colours_used;
	uint32_t colours_important;
};
#pragma pack(pop)

class BMP
{
public:
	BMP(const char *filename, std::string basename);

	void read(const char *filename);
	void write(const char *filename);
	static void write_side_by_side(const BMP& diff, const BMP& base, const BMP& target, const char *filename);

	const std::vector<uint8_t> &get_data() const { return m_data; }
	const std::vector<bool> &get_blurred_edge_mask() const { return m_blurred_edge_mask; }
	int get_width() const { return m_info_header.width; }
	int get_height() const { return m_info_header.height; }
	int get_red_count() const { return m_red_count; }
	int get_yellow_count() const { return m_yellow_count; }
	int get_background_value() const { return m_background_value; }
	int get_non_background_count() const { return m_non_background_count; }

	void increment_red_count(int new_red) { m_red_count += new_red; }
	void increment_yellow_count(int new_yellow) { m_yellow_count += new_yellow; }
	void set_data(std::vector<uint8_t> &new_data);

private:
	int get_average_colour() const;
	int get_non_background_pixel_count(int background_value) const;

	template<int Threshold>
	std::vector<bool> sobel_edges();
	static std::array<int, 2> get_sobel_gradients(int y, int x, const std::vector<uint8_t> &data, int width, int pixel_stride);

	std::vector<bool> blur_edge_mask(const std::vector<bool> &edge_map);

	template<int Radius> // compile-time constant
	static void blur_pixels(int x, int y, int width, int height, std::vector<bool> &mask);

	BMPFileHeader m_file_header;
	BMPInfoHeader m_info_header;

	std::vector<uint8_t> m_data;
	std::vector<bool> m_blurred_edge_mask;
	int m_red_count = 0;
	int m_yellow_count = 0;
	int m_background_value = 0; // used to determine background colour
	int m_non_background_count = 0;
};
#endif

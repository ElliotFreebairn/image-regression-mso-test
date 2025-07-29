//
//
// Copyright the mso-test contributors
//
// SPDX-License-Identifier: MPL-2.0
//
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at http://mozilla.org/MPL/2.0/.

#include <algorithm>
#include <cassert>
#include <cstdlib>
#include <fstream>
#include <iostream>
#include <vector>

#include "bmp.hpp"
#include "pixelbasher.hpp"

struct ParsedArguments
{
	std::string basename;
	std::string extension;
	std::string import_dir;
	std::string export_dir;
	std::string import_compare_dir;
	std::string export_compare_dir;
	std::string image_dump_dir;
	std::string stamp_dir;
	std::vector<BMP> ms_orig_images;
	std::vector<BMP> lo_images;
	std::vector<BMP> ms_conv_images;
	std::vector<BMP> lo_previous_images;
	std::vector<BMP> ms_conv_previous_images;
	bool enable_minor_differences;
	bool image_dump;
	bool lo_previous;
	bool ms_previous;
};

void write_stats_to_csv(const BMP &image1, const BMP &image2, const BMP &diff, int page_number, std::string basename,
						std::string filename)
{
	std::ofstream csv_file(filename);
	if (!csv_file.is_open())
	{
		throw std::runtime_error("Cannot open CSV file for writing");
	}

	double image1_total_pixels = image1.get_width() * image1.get_height();
	double image2_total_pixels = image2.get_width() * image2.get_height();

	csv_file << basename << ","
			 << page_number << ","
			 << image1_total_pixels << ","
			 << image1.get_non_background_count() << ","
			 << (static_cast<double>(image1.get_non_background_count()) / image1_total_pixels) << ","
			 << image2_total_pixels << ","
			 << image2.get_non_background_count() << ","
			 << (static_cast<double>(image2.get_non_background_count()) / image2_total_pixels) << ","
			 << diff.get_red_count() << ","
			 << (static_cast<double>(diff.get_red_count()) / image2_total_pixels) << "\n";
}

void parse_flag(char *argv[], int &arg_index, bool &option, std::string option_name)
{
	std::string value = argv[arg_index - 1];
	if (value == "true" || value == "false")
	{
		option = (value == "true");
		arg_index--;
	}
	else
	{
		throw std::runtime_error("Incorrect usage for " + option_name + ": " + value + " should be true of false");
	}
}

void parse_flags(int &arg_index, char *argv[], ParsedArguments &args)
{
	assert(arg_index >= 4);
	parse_flag(argv, arg_index, args.enable_minor_differences, "minor-differences");
	parse_flag(argv, arg_index, args.image_dump, "image_dump");
	parse_flag(argv, arg_index, args.ms_previous, "ms_previous");
	parse_flag(argv, arg_index, args.lo_previous, "lo_previous");
}

void parse_directories(int &arg_index, char *argv[], ParsedArguments &args)
{
	assert(arg_index >= 6);
	args.stamp_dir = argv[--arg_index];
	args.image_dump_dir = argv[--arg_index];
	args.export_compare_dir = argv[--arg_index];
	args.import_compare_dir = argv[--arg_index];
	args.export_dir = argv[--arg_index];
	args.import_dir = argv[--arg_index];
}

void parse_image_group(char *argv[], int start, int pages, std::vector<BMP> &images, const std::string &basename)
{
	for (int i = 0; i < pages; i++)
	{
		images.push_back(BMP(argv[start + i], basename));
	}
}

ParsedArguments parse_arguments(int argc, char *argv[], int pdf_count = 3)
{
	if (argc < 11)
	{
		throw std::runtime_error("Incorrect usage: " + std::string(argv[0]) + " filename.ext" +
								 "ms_orig-1.bmp ms_orig-2.bmp ... lo-1.bmp lo-2.bmp ... ms_conv.bmp-1.bmp ms_conv-2.bmp ..." +
								 "[lo_previous-1.bmp lo_previous-2.bmp ... ms_conv_previous-1.bmp ms_conv_previous-2.bmp ...]" +
								 "import_dir/ exported_dir/ import-compare_dir/ export-compare_dir/ image-dump_dir/ stamp_dir/" +
								 "[lo_previous] [ms_preivous] [image_dump] [enable_minor_differences]");
	}

	ParsedArguments args;
	int arg_index = argc;

	parse_flags(arg_index, argv, args);
	parse_directories(arg_index, argv, args);

	args.basename = argv[1];
	size_t dot_pos = args.basename.find_last_of('.');
	args.extension = (dot_pos != std::string::npos) ? args.basename.substr(dot_pos + 1) : "";

	if (args.lo_previous)
		pdf_count++;
	if (args.ms_previous)
		pdf_count++;

	int num_image_args = arg_index - 2; // exclude program, basename
	if (num_image_args % pdf_count != 0)
	{
		throw std::runtime_error("Error: " + std::to_string(pdf_count) + " files cannot have a total of " + std::to_string(num_image_args) + " pages.");
	}

	int num_pages = num_image_args / pdf_count;
	int offset = 2;

	parse_image_group(argv, offset, num_pages, args.ms_orig_images, args.basename);
	offset += num_pages;

	parse_image_group(argv, offset, num_pages, args.lo_images, args.basename);
	offset += num_pages;

	parse_image_group(argv, offset, num_pages, args.ms_conv_images, args.basename);
	offset += num_pages;

	if (args.lo_previous)
	{
		parse_image_group(argv, offset, num_pages, args.lo_previous_images, args.basename);
		offset += num_pages;
	}

	if (args.ms_previous)
	{
		parse_image_group(argv, offset, num_pages, args.ms_conv_previous_images, args.basename);
		offset += num_pages;
	}
	return args;
}

BMP diff_and_write(PixelBasher &pixel_basher, BMP &base, BMP &target, bool allow_minor_diffs, const std::string &output_path)
{
	BMP diff = pixel_basher.compare_bmps(base, target, allow_minor_diffs);
	diff.write(output_path.c_str());
	return diff;
}

void compare_lo_previous(PixelBasher &pixelbasher, BMP &base, BMP &lo_diff, BMP &lo_previous,
						 size_t page_index, const ParsedArguments &args)
{
	std::string lo_prev_diff_path = args.import_dir + "/" + args.basename + "_prev-import-page-" + std::to_string(page_index + 1) + ".bmp";
	BMP lo_prev_diff = pixelbasher.compare_bmps(base, lo_previous, args.enable_minor_differences);
	lo_prev_diff.write(lo_prev_diff_path.c_str());

	std::string compare_path = args.import_compare_dir + "/" + args.basename + "_import-compare-page-" + std::to_string(page_index + 1) + ".bmp";
	BMP lo_version_diff = pixelbasher.compare_regressions(base, lo_diff, lo_prev_diff);
	lo_version_diff.write(compare_path.c_str());
}

void compare_ms_previous(PixelBasher &pixelbasher, BMP &base, BMP &ms_conv_diff, BMP &ms_conv_previous,
						 size_t page_index, const ParsedArguments &args)
{
    std::string prev_diff_path = args.export_dir + "/" + args.basename + "_prev-export-page-" + std::to_string(page_index + 1) + ".bmp";
	BMP ms_prev_diff = pixelbasher.compare_bmps(base, ms_conv_previous, args.enable_minor_differences);
	ms_prev_diff.write(prev_diff_path.c_str());

	std::string compare_path = args.export_compare_dir + "/" + args.basename + "_import-compare-page-" + std::to_string(page_index + 1) + ".bmp";
	BMP version_diff = pixelbasher.compare_regressions(base, ms_conv_diff, ms_prev_diff);
	version_diff.write(compare_path.c_str());
}

int main(int argc, char *argv[])
{
	try
	{
        ParsedArguments args = parse_arguments(argc, argv);
		PixelBasher pixel_basher;
		const std::string csv_filename = "diff-pdf-" + args.extension;

		size_t num_pages = args.ms_orig_images.size();
		if (num_pages != args.lo_images.size() || num_pages != args.ms_conv_images.size())
		{
			throw std::runtime_error("Error: mismatched number of pages (" + std::to_string(num_pages) + ") between MS_ORIG, LO, MS_CONV and/or LO_PREVIOUS, MS_CONV_PREVIOUS");
		}

		for (size_t i = 0; i < num_pages; i++)
		{
			BMP base = args.ms_orig_images[i];
			BMP lo = args.lo_images[i];
			BMP ms_conv = args.ms_conv_images[i];

			std::string base_lo_diff_path = args.import_dir + "/" + args.basename + "_import-page-" + std::to_string(i + 1) + ".bmp";
			BMP lo_diff = diff_and_write(pixel_basher, base, lo, args.enable_minor_differences, base_lo_diff_path);
			if (args.image_dump)
			{
				std::string side_by_side_path = args.image_dump_dir + "/lo_comparison-page-" + std::to_string(i + 1) + ".bmp";
				BMP::write_side_by_side(lo_diff, base, lo, args.stamp_dir, side_by_side_path.c_str());
			}

			std::string base_ms_conv_diff_path = args.export_dir + "/" + args.basename + "_export-page-" + std::to_string(i + 1) + ".bmp";
			BMP ms_conv_diff = diff_and_write(pixel_basher, base, ms_conv, args.enable_minor_differences, base_ms_conv_diff_path);
			if (args.image_dump)
			{
				std::string side_by_side_path = args.image_dump_dir + "/ms_conv_comparison-page-" + std::to_string(i + 1) + ".bmp";
				BMP::write_side_by_side(ms_conv_diff, base, ms_conv, args.stamp_dir, side_by_side_path.c_str());
			}

			if (args.lo_previous)
			{
				compare_lo_previous(pixel_basher, base, lo_diff, args.lo_previous_images[i], i, args);
			}

			if (args.ms_previous)
			{
				compare_ms_previous(pixel_basher, base, ms_conv_diff, args.ms_conv_previous_images[i], i, args);
			}

			write_stats_to_csv(base, lo, lo_diff, i + 1, args.basename, (csv_filename + "-import-statistics.csv"));
			write_stats_to_csv(base, ms_conv, ms_conv_diff, i + 1, args.basename, (csv_filename + "-export-statistics.csv"));
		}
	}
	catch (const std::exception &e)
	{
		std::cerr << "Error: " << e.what() << std::endl;
		return 1;
	}
	return 0;
}

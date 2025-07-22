//
//
// Copyright the mso-test contributors
//
// SPDX-License-Identifier: MPL-2.0
//
// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at http://mozilla.org/MPL/2.0/.

#include <iostream>
#include <algorithm>
#include <cstdlib>
#include <fstream>
#include <vector>

#include "bmp.hpp"
#include "pixelbasher.hpp"

struct ParsedArguments {
	std::string basename;
	std::string extension;
	std::string import_dir;
	std::string export_dir;
	bool enable_minor_differences;
	std::vector<BMP> authoritative_images;
	std::vector<BMP> import_images;
	std::vector<BMP> export_images;
};

ParsedArguments parse_arguments(int argc, char* argv[], int pdf_count) {
	if (argc < 7) {
		throw std::runtime_error("Incorrect usage: " + std::string(argv[0]) + " filename.ext auth1.bmp auth2.bmp ... import1.bmp import2.bmp ... export1.bmp export2.bmp import_dir/ exported_dir/ [enable_minor_differences]");
	}

	ParsedArguments args;
	int arg_index = argc;

	std::string last_arg = argv[arg_index - 1];
	if (last_arg == "true" || last_arg == "false") {
		args.enable_minor_differences = (last_arg == "true");
		arg_index--; // Decrease the argument count to exclude the flag
	}

	args.export_dir = argv[arg_index - 1];
	args.import_dir = argv[arg_index - 2];

	args.basename = argv[1]; // The first argument is the basename of the file, not used in this context
	size_t dot_pos = args.basename.find_last_of('.');
	args.extension = (dot_pos != std::string::npos) ? args.basename.substr(dot_pos + 1) : "";

	int num_image_args = arg_index - 4; // from the count and the program name and basename(filename.doc)
	if (num_image_args % pdf_count != 0) {
		throw std::runtime_error("Error: Mismatched number of authoritative and import pages.");
	}

	int num_pages = num_image_args / pdf_count;
	for (int i = 0; i < num_pages; i++)
		args.authoritative_images.push_back(BMP(argv[i + 2], args.basename)); // Plus one to skip the program name

	for (int i = 0; i < num_pages; i++)
		args.import_images.push_back(BMP(argv[i + 2 + num_pages], args.basename)); // Import pages start after the authoritative pages

	for (int i = 0; i < num_pages; i++)
		args.export_images.push_back(BMP(argv[i + 2 + num_pages * 2], args.basename)); // Export pages start after the import page

	return args;
}

void write_stats_to_csv(const BMP& auth, const BMP& lo, const BMP& diff, int page_number, std::string basename,
     std::string filename)
{
  std::ofstream csv_file(filename);
  if (!csv_file.is_open()) {
    throw std::runtime_error("Cannot open CSV file for writing");
  }

  double auth_total_pixels = auth.get_width() * auth.get_height();
  double import_total_pixels = lo.get_width() * lo.get_height();

  csv_file << basename << ","
            << page_number << ","
            << auth_total_pixels << "," // total pixels in auth
            << auth.get_non_background_count() << "," // non-background pixels in auth
            << (static_cast<double>(auth.get_non_background_count()) / auth_total_pixels) << ","
            << import_total_pixels << "," // total pixels in import
            << lo.get_non_background_count() << "," // non-background pixels in import
            << (static_cast<double>(lo.get_non_background_count()) / import_total_pixels)<< ","
            << diff.get_red_count() << "," // red pixels in diff
            << (static_cast<double>(diff.get_red_count()) / import_total_pixels) << "\n";
}

// Main function to compare two BMP images and generate a diff image
// Usage: pixelbasher base.bmp input.bmp output.bmp [enable_minor_differences]
// The last argument is optional and can be "true" or "false" to enable
int main(int argc, char* argv[])
{
	try {
		ParsedArguments args = parse_arguments(argc, argv, 3);
		PixelBasher pixel_basher;

		size_t num_pages = args.authoritative_images.size();
		if (num_pages != args.import_images.size() || num_pages != args.export_images.size()) {
			throw std::runtime_error("Error: Mismatched number of authoritative, import, and export pages.");
		}

		for (size_t i = 0; i < num_pages; i++)  {
			BMP base = args.authoritative_images[i];
			BMP imported = args.import_images[i];
			BMP& export_lo = args.export_images[i];

			// Diffing the authoritative page with the imported page
			std::string import_diff_path = args.import_dir + "/" + args.basename + "_import-page-" + std::to_string(i + 1) + ".bmp";
			BMP diff_result = pixel_basher.compare_to_bmp(base, imported, args.enable_minor_differences);
			diff_result.write(import_diff_path.c_str());

			// Diffing the authoritative page with the exported page
			std::string export_output_path = args.export_dir + "/" + args.basename + "_export-page-" + std::to_string(i + 1) + ".bmp";
			BMP export_diff_result = pixel_basher.compare_to_bmp(base, export_lo, args.enable_minor_differences);
			export_diff_result.write(export_output_path.c_str());

			// Write stats to CSV
			std::string csv_filename = "diff-pdf-" + args.extension;
			write_stats_to_csv(base, imported, diff_result, i + 1, args.basename, (csv_filename + "-import-statistics.csv"));
			write_stats_to_csv(base, export_lo, export_diff_result, i + 1, args.basename, (csv_filename + "-export-statistics.csv"));
		}
    } catch (const std::exception& e) {
		std::cerr << "Error: " << e.what() << std::endl;
		return 1;
	}
	return 0;
}

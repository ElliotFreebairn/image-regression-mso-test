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
	std::string import_compare_dir;
	std::string export_compare_dir;
	bool enable_minor_differences;
	bool lo_previous;
	bool ms_previous;
	std::vector<BMP> authoritative_images;
	std::vector<BMP> import_images;
	std::vector<BMP> export_images;
	std::vector<BMP> import_previous_images;
	std::vector<BMP> export_previous_images;
};

ParsedArguments parse_arguments(int argc, char* argv[], int pdf_count = 3) {
	if (argc < 10) {
		throw std::runtime_error("Incorrect usage: " + std::string(argv[0]) + " filename.ext auth1.bmp auth2.bmp ... import1.bmp import2.bmp ... export1.bmp export2.bmp"  +
		"import_dir/ exported_dir/ import-compare_dir/ export-compare_dir/ [lo_previous] [ms_preivous] [enable_minor_differences]");

	}
	ParsedArguments args;
	int arg_index = argc;

	// Deals with [enable_minor_differences]
	std::string last_arg = argv[arg_index - 1];
	if (last_arg == "true" || last_arg == "false") {
		args.enable_minor_differences = (last_arg == "true");
		arg_index--;
	} else {
		throw std::runtime_error("Incorrect usage: should either be true of false enabling minor differences");
	}

	// Deals with [ms_previous]
	last_arg = argv[arg_index - 1];
	if (last_arg == "true" || last_arg == "false") {
		args.ms_previous = (last_arg == "true");
		arg_index--;
	} else {
		throw std::runtime_error("Incorrect usage: should either be true of false for including ms_previous files");
	}

	// Deals with [ms_preivous]
	last_arg = argv[arg_index - 1];
	if (last_arg == "true" || last_arg == "false") {
		args.lo_previous = (last_arg == "true");
		arg_index--;
	} else {
		throw std::runtime_error("Incorrect usage: should either be true of false for including ms_previous files");
	}

	args.export_compare_dir = argv[arg_index - 1];
	args.import_compare_dir = argv[arg_index - 2];
	args.export_dir = argv[arg_index - 3];
	args.import_dir = argv[arg_index - 4];


	args.basename = argv[1]; // The first argument is the basename of the file, not used in this context
	size_t dot_pos = args.basename.find_last_of('.');
	args.extension = (dot_pos != std::string::npos) ? args.basename.substr(dot_pos + 1) : "";

	if (args.lo_previous) {
		pdf_count++;
	}
	if (args.ms_previous) {
		pdf_count++;
	}

	int num_image_args = arg_index - 6; // exclude program, basename, import, export, import_compare, export_compare dir
	if (num_image_args % pdf_count != 0) {
		throw std::runtime_error("Error: Mismatched number of authoritative and import pages.");
	}

	int num_pages = num_image_args / pdf_count;
	for (int i = 0; i < num_pages; i++) {
		args.authoritative_images.push_back(BMP(argv[i + 2], args.basename)); // Plus one to skip the program name
	}

	for (int i = 0; i < num_pages; i++) {
		args.import_images.push_back(BMP(argv[i + 2 + num_pages], args.basename)); // Import pages start after the authoritative pages
	}

	for (int i = 0; i < num_pages; i++) {
		args.export_images.push_back(BMP(argv[i + 2 + num_pages * 2], args.basename)); // Export pages start after the import page
	}

	if (args.lo_previous) {
		for (int i = 0; i < num_pages; i++) {
			args.import_previous_images.push_back(BMP(argv[i + 2 + num_pages * 3], args.basename));
		}
	}

	if (args.ms_previous) {
		for (int i = 0; i < num_pages; i++) {
			args.export_previous_images.push_back(BMP(argv[i + 2 + num_pages * 4], args.basename));
		}
	}
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
		ParsedArguments args = parse_arguments(argc, argv);
		PixelBasher pixel_basher;

		size_t num_pages = args.authoritative_images.size();
		if (num_pages != args.import_images.size() || num_pages != args.export_images.size()) {
			throw std::runtime_error("Error: Mismatched number of authoritative, import, and export pages.");
		}

		for (size_t i = 0; i < num_pages; i++)  {
			BMP base = args.authoritative_images[i];
			BMP imported = args.import_images[i];
			BMP export_lo = args.export_images[i];

			// Diffing the authoritative page with the imported page
			std::string import_diff_path = args.import_dir + "/" + args.basename + "_import-page-" + std::to_string(i + 1) + ".bmp";
			BMP import_diff_result = pixel_basher.compare_to_bmp(base, imported, args.enable_minor_differences);
			import_diff_result.write(import_diff_path.c_str());

			// Diffing the authoritative page with the exported page
			std::string export_output_path = args.export_dir + "/" + args.basename + "_export-page-" + std::to_string(i + 1) + ".bmp";
			BMP export_diff_result = pixel_basher.compare_to_bmp(base, export_lo, args.enable_minor_differences);
			export_diff_result.write(export_output_path.c_str());

			if (args.lo_previous) {
				BMP import_previous = args.import_previous_images[i];

				std::string import_previous_output_path = args.import_dir + "/"  + args.basename + "_prev-import-page-" + std::to_string(i + 1) + ".bmp";
				BMP import_previous_result = pixel_basher.compare_to_bmp(base, import_previous, args.enable_minor_differences);
				import_previous_result.write(import_previous_output_path.c_str());

				std::string import_previous_comapare_output_path = args.import_compare_dir + "/" + args.basename + "_import-compare-page-" + std::to_string(i + 1) + ".bmp";
				BMP current_previous_result = pixel_basher.compare_regressions(base, import_diff_result, import_previous_result);
				current_previous_result.write(import_previous_comapare_output_path.c_str());
			}

			if (args.ms_previous) {
				BMP export_previous = args.export_previous_images[i];

				std::string export_previous_output_path = args.export_dir + "/"  + args.basename + "_prev-export-page-" + std::to_string(i + 1) + ".bmp";
				BMP export_previous_result = pixel_basher.compare_to_bmp(base, export_previous, args.enable_minor_differences);
				export_previous_result.write(export_previous_output_path.c_str());

				std::string export_previous_compare_output_path = args.export_compare_dir + "/" + args.basename + "_import-compare-page-" + std::to_string(i + 1) + ".bmp";
				BMP current_previous_result = pixel_basher.compare_regressions(base, export_diff_result, export_previous_result);
				current_previous_result.write(export_previous_compare_output_path.c_str());
			}

			// Write stats to CSV
			std::string csv_filename = "diff-pdf-" + args.extension;
			write_stats_to_csv(base, imported, import_diff_result, i + 1, args.basename, (csv_filename + "-import-statistics.csv"));
			write_stats_to_csv(base, export_lo, export_diff_result, i + 1, args.basename, (csv_filename + "-export-statistics.csv"));
		}
    } catch (const std::exception& e) {
		std::cerr << "Error: " << e.what() << std::endl;
		return 1;
	}
	return 0;
}

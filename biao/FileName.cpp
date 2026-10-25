#include <iostream>
#include <string>
#include <vector>
#include <map>
#include <cmath>
#include <algorithm>
#include <filesystem>
#include <sstream>
#include <iomanip>
#include <xlsxio_read.h>
#include <xlsxio_write.h>

namespace fs = std::filesystem;

int main(int argc, char* argv[]) {
    try {
        // 1. Check input
        if (argc < 2) {
            std::cerr << "=========================================" << std::endl;
            std::cerr << "Usage: Drag and drop an .xlsx file onto this program." << std::endl;
            std::cerr << "=========================================" << std::endl;
            system("pause");
            return 1;
        }
        std::string input_filename = argv[1];

        // 2. Open file and sheet
        xlsxioreader read_handle = xlsxioread_open(input_filename.c_str());
        if (!read_handle) {
            std::cerr << "Failed to open input file." << std::endl;
            system("pause");
            return 1;
        }

        xlsxioreadersheet sheet_handle = xlsxioread_sheet_open(read_handle, NULL, XLSXIOREAD_SKIP_NONE);
        if (!sheet_handle) {
            std::cerr << "Failed to open sheet." << std::endl;
            xlsxioread_close(read_handle);
            system("pause");
            return 1;
        }

        // 3. Read all rows into memory
        std::vector<std::vector<std::string>> all_rows;
        const char* cell_value = nullptr;

        while (xlsxioread_sheet_next_row(sheet_handle)) {
            std::vector<std::string> row_cells;
            while ((cell_value = xlsxioread_sheet_next_cell(sheet_handle)) != nullptr) {
                row_cells.push_back(cell_value ? cell_value : "");
            }
            bool isEmpty = true;
            for (const auto& c : row_cells) {
                if (!c.empty()) { isEmpty = false; break; }
            }
            if (!isEmpty) {
                all_rows.push_back(row_cells);
            }
        }

        xlsxioread_sheet_close(sheet_handle);
        xlsxioread_close(read_handle);

        if (all_rows.size() < 2) {
            std::cerr << "Not enough data rows to parse." << std::endl;
            system("pause");
            return 1;
        }

        // 4. 核心重构：智能定位单价行和表头
        std::vector<std::string> headers;
        std::vector<double> prices;
        int data_start_index = 0;
        size_t col_offset = 1; // 跳过第一列的“种类/序号”

        // 扫描第一列，寻找包含“单价”的行
        int price_row_idx = -1;
        for (size_t r = 0; r < all_rows.size(); ++r) {
            if (!all_rows[r].empty() && all_rows[r][0].find("单价") != std::string::npos) {
                price_row_idx = r;
                break;
            }
        }

        if (price_row_idx != -1) {
            // 如果找到了单价行，上一行就是表头
            int header_row_idx = (price_row_idx > 0) ? price_row_idx - 1 : -1;

            if (header_row_idx >= 0) {
                for (size_t i = col_offset; i < all_rows[header_row_idx].size(); ++i) {
                    std::string h = all_rows[header_row_idx][i];
                    headers.push_back(h.empty() ? ("Item_" + std::to_string(i)) : h);
                }
            }
            else {
                for (size_t i = col_offset; i < all_rows[price_row_idx].size(); ++i) {
                    headers.push_back("Item_" + std::to_string(i));
                }
            }

            // 读取单价
            for (size_t i = col_offset; i < all_rows[price_row_idx].size(); ++i) {
                try { prices.push_back(std::stod(all_rows[price_row_idx][i])); }
                catch (...) { prices.push_back(0.0); }
            }
            data_start_index = price_row_idx + 1; // 数据从单价行的下一行开始
        }
        else {
            // 备用方案：如果没有“单价”行，则使用旧逻辑
            bool hasHeader = false;
            int numCount = 0, totalCount = 0;
            for (size_t i = 1; i < all_rows[0].size(); ++i) {
                totalCount++;
                try {
                    if (!all_rows[0][i].empty()) { std::stod(all_rows[0][i]); numCount++; }
                }
                catch (...) {}
            }
            if (numCount <= totalCount / 2) hasHeader = true;

            if (hasHeader) {
                for (size_t i = col_offset; i < all_rows[0].size(); ++i) headers.push_back(all_rows[0][i]);
                for (size_t i = col_offset; i < all_rows[1].size(); ++i) {
                    try { prices.push_back(std::stod(all_rows[1][i])); }
                    catch (...) { prices.push_back(0.0); }
                }
                data_start_index = 2;
            }
            else {
                for (size_t i = col_offset; i < all_rows[0].size(); ++i) headers.push_back("Item_" + std::to_string(i));
                for (size_t i = col_offset; i < all_rows[0].size(); ++i) {
                    try { prices.push_back(std::stod(all_rows[0][i])); }
                    catch (...) { prices.push_back(0.0); }
                }
                data_start_index = 1;
            }
        }

        std::map<std::string, double> priceMap;
        for (size_t i = 0; i < headers.size() && i < prices.size(); ++i) {
            priceMap[headers[i]] = prices[i];
        }

        // 5. 统计人名及购买数量
        std::map<std::string, std::map<std::string, int>> stats;
        for (size_t r = data_start_index; r < all_rows.size(); ++r) {
            const auto& row = all_rows[r];
            for (size_t i = 0; i < headers.size() && (i + col_offset) < row.size(); ++i) {
                std::string person = row[i + col_offset];
                if (!person.empty()) {
                    stats[person][headers[i]]++;
                }
            }
        }

        // 6. 生成输出 Excel
        fs::path input_path(input_filename);
        fs::path output_path = input_path.parent_path() / "result.xlsx";
        std::string output_filename = output_path.string();

        xlsxiowriter write_handle = xlsxiowrite_open(output_filename.c_str(), "Sheet1");
        if (!write_handle) {
            std::cerr << "Failed to create output file." << std::endl;
            system("pause");
            return 1;
        }

        xlsxiowrite_next_row(write_handle);
        xlsxiowrite_add_cell_string(write_handle, "Name");
        xlsxiowrite_add_cell_string(write_handle, "Items");
        xlsxiowrite_add_cell_string(write_handle, "Total");

        for (const auto& pair_person : stats) {
            const std::string& person = pair_person.first;
            const std::map<std::string, int>& items = pair_person.second;

            double total = 0.0;
            std::string detail_str;

            for (const auto& pair_item : items) {
                const std::string& good = pair_item.first;
                int cnt = pair_item.second;
                total += priceMap[good] * cnt;
                if (!detail_str.empty()) detail_str += ", ";
                detail_str += good + " x" + std::to_string(cnt);
            }

            total = std::round(total * 100.0) / 100.0;
            std::stringstream ss;
            ss << std::fixed << std::setprecision(2) << total;
            std::string total_str = ss.str();

            xlsxiowrite_next_row(write_handle);
            xlsxiowrite_add_cell_string(write_handle, person.c_str());
            xlsxiowrite_add_cell_string(write_handle, detail_str.c_str());
            xlsxiowrite_add_cell_string(write_handle, total_str.c_str());
        }

        xlsxiowrite_close(write_handle);

        std::cout << "=========================================" << std::endl;
        std::cout << "Success! Output file: " << fs::absolute(output_filename).string() << std::endl;
        std::cout << "=========================================" << std::endl;
        system("pause");
    }
    catch (const std::exception& e) {
        std::cerr << "Runtime Error: " << e.what() << std::endl;
        system("pause");
        return 1;
    }
    catch (...) {
        std::cerr << "Unknown Error!" << std::endl;
        system("pause");
        return 1;
    }
    return 0;
}
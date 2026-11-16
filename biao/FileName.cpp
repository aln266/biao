#define NOMINMAX 
#define _CRT_SECURE_NO_WARNINGS
#pragma execution_character_set("utf-8")

#include <iostream>
#include <string>
#include <vector>
#include <map>
#include <cmath>
#include <algorithm>
#include <filesystem>
#include <sstream>
#include <iomanip>
#include <fstream>
#include <windows.h>
#include <shellapi.h> 
#include <xlsxio_read.h>
#include <xlsxio_write.h>

namespace fs = std::filesystem;

// ================= 数据结构 =================
struct TableData {
    std::vector<std::string> headers;
    std::map<std::string, double> prices;
    std::map<std::string, std::map<std::string, int>> stats;
};

// 全局数据存储器
std::map<std::string, std::map<std::string, int>> global_stats;
std::map<std::string, double> global_prices;
std::map<std::string, std::string> global_status;
std::map<std::string, std::map<std::string, std::string>> user_goods_status;

std::string exe_dir;
std::string current_store_path = "";
std::string current_store_name = "";

// ================= 编码转换核心 (仅用于文件路径) =================
std::string UTF8ToGBK(const std::string& str) {
    if (str.empty()) return "";
    int len = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, NULL, 0);
    wchar_t* wstr = new wchar_t[len];
    MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, wstr, len);
    len = WideCharToMultiByte(CP_ACP, 0, wstr, -1, NULL, 0, NULL, NULL);
    char* gbk_str = new char[len];
    WideCharToMultiByte(CP_ACP, 0, wstr, -1, gbk_str, len, NULL, NULL);
    std::string res(gbk_str);
    delete[] wstr; delete[] gbk_str;
    return res;
}

std::string WStringToUTF8(const std::wstring& wstr) {
    if (wstr.empty()) return "";
    int len = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, NULL, 0, NULL, NULL);
    char* utf8_str = new char[len];
    WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, utf8_str, len, NULL, NULL);
    std::string res(utf8_str);
    delete[] utf8_str;
    return res;
}

std::string get_exe_dir() {
    wchar_t buffer[MAX_PATH];
    GetModuleFileNameW(NULL, buffer, MAX_PATH);
    std::string res = WStringToUTF8(buffer);
    std::string::size_type pos = res.find_last_of("\\/");
    return res.substr(0, pos);
}

std::string extract_stem(const std::string& filepath) {
    size_t pos = filepath.find_last_of("\\/");
    std::string filename = (pos != std::string::npos) ? filepath.substr(pos + 1) : filepath;
    size_t dot_pos = filename.find_last_of('.');
    return (dot_pos != std::string::npos) ? filename.substr(0, dot_pos) : filename;
}

std::string clean_path(std::string path) {
    path.erase(0, path.find_first_not_of(" \t\r\n"));
    path.erase(path.find_last_not_of(" \t\r\n") + 1);
    if (path.length() >= 2) {
        if ((path.front() == '"' && path.back() == '"') || (path.front() == '“' && path.back() == '”')) {
            path = path.substr(1, path.length() - 2);
        }
    }
    return path;
}
// =================================================================

void clear_global_data() {
    global_stats.clear();
    global_prices.clear();
    global_status.clear();
    user_goods_status.clear();
}

// ================= 本地状态数据读写 =================
void load_store_status() {
    global_status.clear();
    user_goods_status.clear();

    std::string status_file = UTF8ToGBK(current_store_path + "\\global_status.txt");
    std::ifstream fin(status_file);
    if (fin.is_open()) {
        std::string line;
        while (std::getline(fin, line)) {
            if (line.empty()) continue;
            size_t pos = line.find('|');
            if (pos != std::string::npos) global_status[line.substr(0, pos)] = line.substr(pos + 1);
        }
    }

    std::string user_status_file = UTF8ToGBK(current_store_path + "\\user_goods_status.txt");
    std::ifstream fin2(user_status_file);
    if (fin2.is_open()) {
        std::string line;
        while (std::getline(fin2, line)) {
            if (line.empty()) continue;
            size_t p1 = line.find('|'); size_t p2 = line.find('|', p1 + 1);
            if (p1 != std::string::npos && p2 != std::string::npos) {
                user_goods_status[line.substr(0, p1)][line.substr(p1 + 1, p2 - p1 - 1)] = line.substr(p2 + 1);
            }
        }
    }
}

void save_store_status() {
    std::string status_file = UTF8ToGBK(current_store_path + "\\global_status.txt");
    std::ofstream fout(status_file);
    for (const auto& p : global_status) fout << p.first << "|" << p.second << "\n";
    fout.close();

    std::string user_status_file = UTF8ToGBK(current_store_path + "\\user_goods_status.txt");
    std::ofstream fout2(user_status_file);
    for (const auto& person : user_goods_status) {
        for (const auto& good : person.second) {
            fout2 << person.first << "|" << good.first << "|" << good.second << "\n";
        }
    }
}

// ================= 批次文件保存与全局聚合 =================
void save_batch_data(const std::string& batch_name, const std::map<std::string, double>& prices, const std::map<std::string, std::map<std::string, int>>& stats) {
    std::string batch_path = UTF8ToGBK(current_store_path + "\\batches\\" + batch_name);
    if (!fs::exists(batch_path)) fs::create_directories(batch_path);

    std::ofstream fout(batch_path + "\\data.txt");
    for (const auto& p : prices) fout << "[PRICE]" << p.first << "|" << p.second << "\n";
    for (const auto& person : stats) {
        for (const auto& item : person.second) {
            fout << "[BUY]" << person.first << "|" << item.first << "|" << item.second << "\n";
        }
    }
}

void aggregate_all_batches() {
    global_stats.clear(); global_prices.clear();
    std::string batches_root = UTF8ToGBK(current_store_path + "\\batches");
    if (!fs::exists(batches_root)) return;

    for (const auto& entry : fs::directory_iterator(batches_root)) {
        if (entry.is_directory()) {
            std::string batch_name_gbk = entry.path().filename().string();
            wchar_t wbuf[MAX_PATH];
            MultiByteToWideChar(CP_ACP, 0, batch_name_gbk.c_str(), -1, wbuf, MAX_PATH);
            std::string batch_name_utf8 = WStringToUTF8(wbuf);

            std::string data_file = entry.path().string() + "\\data.txt";
            std::ifstream fin(data_file);
            if (!fin.is_open()) continue;

            std::map<std::string, double> current_batch_prices;
            std::vector<std::tuple<std::string, std::string, int>> current_batch_buys;

            std::string line;
            while (std::getline(fin, line)) {
                if (line.empty()) continue;
                if (line.substr(0, 7) == "[PRICE]") {
                    std::string content = line.substr(7); size_t pos = content.find('|');
                    if (pos != std::string::npos) current_batch_prices[content.substr(0, pos)] = std::stod(content.substr(pos + 1));
                }
                else if (line.substr(0, 5) == "[BUY]") {
                    std::string content = line.substr(5); size_t p1 = content.find('|'); size_t p2 = content.find('|', p1 + 1);
                    if (p1 != std::string::npos && p2 != std::string::npos) {
                        current_batch_buys.push_back(std::make_tuple(content.substr(0, p1), content.substr(p1 + 1, p2 - p1 - 1), std::stoi(content.substr(p2 + 1))));
                    }
                }
            }

            for (auto& buy : current_batch_buys) {
                std::string person = std::get<0>(buy);
                std::string good = std::get<1>(buy);
                int cnt = std::get<2>(buy);
                std::string combined_good = "[" + batch_name_utf8 + "] " + good;

                global_stats[person][combined_good] += cnt;
                if (current_batch_prices.count(good)) global_prices[combined_good] = current_batch_prices[good];
            }
        }
    }

    for (const auto& person : global_stats) {
        for (const auto& item : person.second) {
            if (user_goods_status[person.first].find(item.first) == user_goods_status[person.first].end()) {
                user_goods_status[person.first][item.first] = "未排发";
            }
        }
    }
    save_store_status();
}

// ================= Excel 解析 =================
TableData parse_excel(const std::string& input_filename) {
    TableData data;
    std::string gbk_path = UTF8ToGBK(input_filename);
    FILE* test_file = fopen(gbk_path.c_str(), "rb");
    if (!test_file) return data;
    fclose(test_file);

    xlsxioreader read_handle = xlsxioread_open(gbk_path.c_str());
    if (!read_handle) return data;
    xlsxioreadersheet sheet_handle = xlsxioread_sheet_open(read_handle, NULL, XLSXIOREAD_SKIP_NONE);
    if (!sheet_handle) { xlsxioread_close(read_handle); return data; }

    std::vector<std::vector<std::string>> all_rows;
    const char* cell_value = nullptr;
    while (xlsxioread_sheet_next_row(sheet_handle)) {
        std::vector<std::string> row_cells;
        while ((cell_value = xlsxioread_sheet_next_cell(sheet_handle)) != nullptr) {
            row_cells.push_back(cell_value ? cell_value : "");
        }
        bool isEmpty = true;
        for (const auto& c : row_cells) { if (!c.empty()) { isEmpty = false; break; } }
        if (!isEmpty) all_rows.push_back(row_cells);
    }
    xlsxioread_sheet_close(sheet_handle); xlsxioread_close(read_handle);
    if (all_rows.size() < 2) return data;

    int price_row_idx = -1, max_num_count = 0;
    size_t scan_limit = std::min(all_rows.size(), (size_t)5);
    for (size_t r = 0; r < scan_limit; ++r) {
        int numCount = 0;
        for (size_t c = 1; c < all_rows[r].size(); ++c) {
            try { if (!all_rows[r][c].empty()) { (void)std::stod(all_rows[r][c]); numCount++; } }
            catch (...) {}
        }
        if (numCount > max_num_count) { max_num_count = numCount; price_row_idx = (int)r; }
    }

    if (price_row_idx != -1 && max_num_count > 0) {
        int data_start_index = price_row_idx + 1;
        size_t col_offset = 1;
        int header_row_idx = price_row_idx - 1;
        size_t max_cols = all_rows[price_row_idx].size();
        if (header_row_idx >= 0 && all_rows[header_row_idx].size() > max_cols) max_cols = all_rows[header_row_idx].size();

        for (size_t i = col_offset; i < max_cols; ++i) {
            std::string h = "";
            if (header_row_idx >= 0 && i < all_rows[header_row_idx].size()) h = all_rows[header_row_idx][i];
            data.headers.push_back(h.empty() ? ("Item_" + std::to_string(i)) : h);
        }

        for (size_t i = col_offset; i < all_rows[price_row_idx].size(); ++i) {
            std::string good_name = data.headers[i - col_offset];
            try { data.prices[good_name] = std::stod(all_rows[price_row_idx][i]); }
            catch (...) { data.prices[good_name] = 0.0; }
        }

        for (size_t r = data_start_index; r < all_rows.size(); ++r) {
            const auto& row = all_rows[r];
            for (size_t i = 0; i < data.headers.size(); ++i) {
                size_t col_idx = i + col_offset;
                if (col_idx < row.size()) {
                    std::string person = row[col_idx];
                    person.erase(0, person.find_first_not_of(" \t\r\n"));
                    person.erase(person.find_last_not_of(" \t\r\n") + 1);
                    if (!person.empty()) data.stats[person][data.headers[i]]++;
                }
            }
        }
    }
    return data;
}

// ================= Excel 导出 =================
// 导出全局汇总表 (Summary.xlsx)
void export_summary() {
    std::string output_filename = current_store_path + "\\Summary.xlsx";
    std::string gbk_out_path = UTF8ToGBK(output_filename);

    xlsxiowriter write_handle = xlsxiowrite_open(gbk_out_path.c_str(), "Sheet1");
    if (!write_handle) return;

    xlsxiowrite_next_row(write_handle);
    xlsxiowrite_add_cell_string(write_handle, "cn");
    xlsxiowrite_add_cell_string(write_handle, "\xE5\x88\xB6\xE5\x93\x81");
    xlsxiowrite_add_cell_string(write_handle, "zfb");
    xlsxiowrite_add_cell_string(write_handle, "wx");

    std::map<std::string, std::map<std::string, int>> export_stats;
    std::map<std::string, double> export_prices;

    for (const auto& pair_person : global_stats) {
        for (const auto& pair_item : pair_person.second) {
            std::string good_with_batch = pair_item.first;
            int cnt = pair_item.second;
            std::string pure_good = good_with_batch;
            size_t pos = pure_good.find("] ");
            if (pos != std::string::npos && pure_good.front() == '[') pure_good = pure_good.substr(pos + 2);

            export_stats[pair_person.first][pure_good] += cnt;
            if (export_prices.find(pure_good) == export_prices.end() && global_prices.count(good_with_batch)) {
                export_prices[pure_good] = global_prices[good_with_batch];
            }
        }
    }

    for (const auto& pair_person : export_stats) {
        double total = 0.0; std::string detail_str;
        for (const auto& pair_item : pair_person.second) {
            std::string good = pair_item.first;
            int cnt = pair_item.second;
            total += (export_prices.count(good) ? export_prices[good] : 0.0) * cnt;
            if (!detail_str.empty()) detail_str += ", ";
            detail_str += good + " x" + std::to_string(cnt);
        }

        total = std::round(total * 100.0) / 100.0;
        int fee_tier = static_cast<int>(std::ceil(total / 100.0));
        double wx = std::round((total + (fee_tier * 0.1)) * 100.0) / 100.0;

        std::stringstream ss_zfb, ss_wx;
        ss_zfb << std::fixed << std::setprecision(2) << total;
        ss_wx << std::fixed << std::setprecision(2) << wx;

        xlsxiowrite_next_row(write_handle);
        xlsxiowrite_add_cell_string(write_handle, pair_person.first.c_str());
        xlsxiowrite_add_cell_string(write_handle, detail_str.c_str());
        xlsxiowrite_add_cell_string(write_handle, ss_zfb.str().c_str());
        xlsxiowrite_add_cell_string(write_handle, ss_wx.str().c_str());
    }
    xlsxiowrite_close(write_handle);
}

// 导出单批次汇总表 (批次汇总.xlsx)
void export_single_batch(const std::string& batch_name, const TableData& data) {
    std::string output_filename = current_store_path + "\\batches\\" + batch_name + "\\批次汇总.xlsx";
    std::string gbk_out_path = UTF8ToGBK(output_filename);

    xlsxiowriter write_handle = xlsxiowrite_open(gbk_out_path.c_str(), "Sheet1");
    if (!write_handle) return;

    xlsxiowrite_next_row(write_handle);
    xlsxiowrite_add_cell_string(write_handle, "cn");
    xlsxiowrite_add_cell_string(write_handle, "\xE5\x88\xB6\xE5\x93\x81");
    xlsxiowrite_add_cell_string(write_handle, "zfb");
    xlsxiowrite_add_cell_string(write_handle, "wx");

    for (const auto& pair_person : data.stats) {
        double total = 0.0; std::string detail_str;
        for (const auto& pair_item : pair_person.second) {
            std::string good = pair_item.first;
            int cnt = pair_item.second;
            double price = data.prices.count(good) ? data.prices.at(good) : 0.0;
            total += price * cnt;
            if (!detail_str.empty()) detail_str += ", ";
            detail_str += good + " x" + std::to_string(cnt);
        }

        total = std::round(total * 100.0) / 100.0;
        int fee_tier = static_cast<int>(std::ceil(total / 100.0));
        double wx = std::round((total + (fee_tier * 0.1)) * 100.0) / 100.0;

        std::stringstream ss_zfb, ss_wx;
        ss_zfb << std::fixed << std::setprecision(2) << total;
        ss_wx << std::fixed << std::setprecision(2) << wx;

        xlsxiowrite_next_row(write_handle);
        xlsxiowrite_add_cell_string(write_handle, pair_person.first.c_str());
        xlsxiowrite_add_cell_string(write_handle, detail_str.c_str());
        xlsxiowrite_add_cell_string(write_handle, ss_zfb.str().c_str());
        xlsxiowrite_add_cell_string(write_handle, ss_wx.str().c_str());
    }
    xlsxiowrite_close(write_handle);
}

bool import_excel_to_batch(const std::string& filepath, const std::string& batch_name) {
    std::cout << "\n正在解析表格: " << filepath << "\n";
    TableData newTable = parse_excel(filepath);
    if (newTable.prices.empty()) { std::cout << "❌ 错误：未能读取到有效数据。\n"; return false; }

    std::string final_batch_name = batch_name;
    std::string batch_path = UTF8ToGBK(current_store_path + "\\batches\\" + final_batch_name);

    if (fs::exists(batch_path)) {
        std::cout << "\n⚠️ 警告：批次【" << final_batch_name << "】已存在！\n1. 覆盖该批次\n2. 追加为新批次\n请选择(1/2): ";
        std::string overwrite; std::getline(std::cin, overwrite);
        if (overwrite == "2") {
            time_t now = time(0); struct tm local_tm; localtime_s(&local_tm, &now);
            char time_buf[80]; strftime(time_buf, sizeof(time_buf), "_%H%M%S", &local_tm);
            final_batch_name = final_batch_name + time_buf;
            std::cout << "已自动重命名为新批次: " << final_batch_name << "\n";
        }
        else {
            fs::remove_all(batch_path);
        }
    }

    // 保存底层数据
    save_batch_data(final_batch_name, newTable.prices, newTable.stats);

    // 🔥 生成该批次独立的 Excel 表格
    export_single_batch(final_batch_name, newTable);

    // 重新聚合全局数据并生成全局表格
    aggregate_all_batches();
    export_summary();

    std::cout << "✅ 批次导入成功！\n";
    std::cout << "  - 已生成批次独立账单: data\\" << current_store_name << "\\batches\\" << final_batch_name << "\\批次汇总.xlsx\n";
    std::cout << "  - 已更新全局汇总表: data\\" << current_store_name << "\\Summary.xlsx\n";
    return true;
}

void admin_page() {
    load_store_status(); aggregate_all_batches();
    while (true) {
        system("cls");
        std::cout << "=========================================\n";
        std::cout << "       【管理员控制台】 当前店铺: " << current_store_name << "\n";
        std::cout << "=========================================\n";
        std::cout << "  1. 导入新批次表格\n  2. 按用户管理商品状态\n  0. 返回主菜单\n=========================================\n请选择: ";

        std::string choice; std::getline(std::cin, choice);

        if (choice == "1") {
            system("cls");
            std::cout << "请拖拽 Excel 文件到窗口内并按回车 (输入 0 取消): ";
            std::string filepath_utf8; std::getline(std::cin, filepath_utf8);
            if (filepath_utf8 == "0" || filepath_utf8.empty()) continue;
            filepath_utf8 = clean_path(filepath_utf8);
            std::string batch_name = extract_stem(filepath_utf8);
            std::cout << "自动识别批次名为: " << batch_name << "\n";
            import_excel_to_batch(filepath_utf8, batch_name);
            system("pause");
        }
        else if (choice == "2") {
            system("cls");
            std::cout << "请输入要管理的【用户姓名】: ";
            std::string input_name; std::getline(std::cin, input_name);
            if (input_name.empty()) continue;
            auto it = global_stats.find(input_name);
            if (it == global_stats.end()) { std::cout << " -> 未找到。按任意键返回...\n"; system("pause > nul"); continue; }

            while (true) {
                system("cls");
                std::cout << "========== [" << input_name << "] 的商品状态管理 ==========\n";
                std::vector<std::string> goods_list; int idx = 1;
                for (const auto& pair_item : it->second) {
                    std::string status = user_goods_status[input_name].count(pair_item.first) ? user_goods_status[input_name][pair_item.first] : "未排发";
                    goods_list.push_back(pair_item.first);
                    std::cout << idx++ << ". " << pair_item.first << " [" << status << "]\n";
                }
                std::cout << "========================================================\n请输入要修改的商品序号（支持多选空格隔开，0返回）: ";
                std::string input; std::getline(std::cin, input);
                if (input == "0" || input.empty()) break;
                std::vector<int> targets; std::stringstream ss(input); int item_idx;
                while (ss >> item_idx) if (item_idx >= 1 && item_idx <= (int)goods_list.size()) targets.push_back(item_idx - 1);

                if (!targets.empty()) {
                    std::cout << "\n请选择新状态：\n1. 未排发\n2. 已排发\n输入序号: ";
                    std::string status_choice; std::getline(std::cin, status_choice);
                    std::string new_status = (status_choice == "1") ? "未排发" : (status_choice == "2" ? "已排发" : "");
                    if (new_status.empty()) continue;
                    for (int t : targets) user_goods_status[input_name][goods_list[t]] = new_status;
                    save_store_status(); export_summary();
                    std::cout << " ✅ 已成功批量修改 " << targets.size() << " 个商品的状态！按任意键继续...\n";
                    system("pause > nul");
                }
            }
        }
        else if (choice == "0") break;
    }
}

void user_page() {
    system("cls");
    std::cout << "=========================================\n           【订单查询系统】\n=========================================\n输入 0 退出并返回主菜单\n\n请输入要查询的【店铺名】: ";
    std::string store; std::getline(std::cin, store);
    if (store == "0") return;

    std::string target_path = exe_dir + "\\data\\" + store;
    if (!fs::exists(UTF8ToGBK(target_path))) { std::cout << " -> 错误：店铺不存在。\n"; system("pause > nul"); return; }
    current_store_path = target_path; current_store_name = store;
    load_store_status(); aggregate_all_batches();

    std::cout << "\n已连接店铺：【" << store << "】\n";
    std::string input_name;
    while (true) {
        std::cout << "请输入姓名 (输入 0 换店铺): ";
        if (!std::getline(std::cin, input_name)) break;
        if (input_name.empty()) continue;
        if (input_name == "0") break;
        auto it = global_stats.find(input_name);
        if (it == global_stats.end()) std::cout << " -> 未找到该用户的记录。\n\n";
        else {
            std::cout << "\n -> 【" << input_name << "】在所有批次中的购买汇总如下：\n";
            for (const auto& pair_item : it->second) {
                std::string status = user_goods_status[input_name].count(pair_item.first) ? user_goods_status[input_name][pair_item.first] : "未排发";
                std::cout << "    - " << pair_item.first << " x" << pair_item.second << "  [" << status << "]\n";
            }
            std::cout << "\n";
        }
    }
}

int main() {
    SetConsoleOutputCP(65001); SetConsoleCP(65001);
    SetConsoleTitleA("多店铺订单管理系统 v3.6");
    exe_dir = get_exe_dir();

    std::string data_root = exe_dir + "\\data";
    if (!fs::exists(UTF8ToGBK(data_root))) fs::create_directories(UTF8ToGBK(data_root));

    int argc_w;
    LPWSTR* argv_w = CommandLineToArgvW(GetCommandLineW(), &argc_w);

    if (argc_w > 1) {
        std::wstring wfilepath = argv_w[1];
        if (wfilepath.length() >= 2 && wfilepath.front() == L'"' && wfilepath.back() == L'"') wfilepath = wfilepath.substr(1, wfilepath.length() - 2);
        std::string filepath_utf8 = WStringToUTF8(wfilepath);

        system("cls");
        std::cout << "=========================================\n      检测到拖拽文件，启动快捷导入\n=========================================\n请输入您的【店铺名】: ";
        std::string store; std::getline(std::cin, store);
        if (!store.empty()) {
            current_store_name = store; current_store_path = data_root + "\\" + store;
            if (!fs::exists(UTF8ToGBK(current_store_path))) fs::create_directories(UTF8ToGBK(current_store_path));
            std::string batch_name = extract_stem(filepath_utf8);
            std::cout << "自动识别批次名为: " << batch_name << "\n";
            load_store_status(); import_excel_to_batch(filepath_utf8, batch_name); system("pause");
            admin_page();
        }
        LocalFree(argv_w); return 0;
    }
    if (argv_w) LocalFree(argv_w);

    while (true) {
        system("cls");
        std::cout << "=========================================\n       欢迎使用 多店铺订单管理系统\n=========================================\n  1. 管理员模式 (创建/进入店铺)\n  2. 普通用户模式 (跨店查询订单)\n  0. 退出系统\n=========================================\n请选择身份 (输入数字): ";
        std::string choice; std::getline(std::cin, choice);
        if (choice == "1") {
            system("cls");
            std::cout << "请输入您的【店铺名】(首次输入将自动创建): ";
            std::string store; std::getline(std::cin, store);
            if (store.empty()) continue;
            current_store_name = store; current_store_path = data_root + "\\" + store;
            if (!fs::exists(UTF8ToGBK(current_store_path))) {
                fs::create_directories(UTF8ToGBK(current_store_path));
                std::cout << "✅ 店铺【" << store << "】创建成功！\n"; system("pause > nul");
            }
            admin_page(); clear_global_data();
        }
        else if (choice == "2") {
            user_page(); clear_global_data();
        }
        else if (choice == "0") break;
    }
    return 0;
}
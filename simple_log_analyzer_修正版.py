#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Log執行時間分析工具 (僅處理當前資料夾，排除.exe和.bat)

這個腳本使用簡單的前綴後綴搜尋方式從log檔案中提取執行時間，
並計算統計數據（平均值、中位數、分位數等），將結果匯出到Excel檔案。

使用方法:
    python simple_log_analyzer_修正版.py log資料夾路徑 前綴字串 後綴字串 [選項]

選項:
    -o, --output FILE          指定輸出的Excel檔案路徑 (默認為資料夾名稱加上"_分析結果.xlsx")
"""

import argparse
import os
import sys
import pandas as pd
import numpy as np
import time
import logging
import glob
import re
from datetime import datetime

# 設置日誌
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)


def parse_args():
    """解析命令行參數"""
    parser = argparse.ArgumentParser(description="Log執行時間分析工具 - 排除.exe和.bat")
    parser.add_argument("log_dir", help="要分析的log檔案資料夾路徑")
    parser.add_argument("prefix", help="要搜尋的前綴字串")
    parser.add_argument("suffix", help="要搜尋的後綴字串")

    return parser.parse_args()


def get_all_files(log_dir):
    """獲取指定資料夾中的所有檔案 (不包含子資料夾，排除.exe和.bat)"""
    all_files = []

    # 只處理當前資料夾中的檔案
    pattern = os.path.join(log_dir, "*")
    files = glob.glob(pattern)

    # 排除的副檔名
    excluded_extensions = {'.exe', '.bat'}

    for f in files:
        if os.path.isfile(f):
            # 取得檔案副檔名
            file_ext = os.path.splitext(f)[1].lower()

            # 如果副檔名不在排除清單中，就包含這個檔案
            if file_ext not in excluded_extensions:
                all_files.append(f)
                logger.debug(f"包含檔案: {os.path.basename(f)}")
            else:
                logger.debug(f"排除檔案: {os.path.basename(f)} (副檔名: {file_ext})")

    return all_files


def extract_times_from_file(log_file, prefix, suffix):
    """從單個log檔案中提取執行時間"""
    try:
        with open(log_file, 'r', encoding='utf-8', errors='ignore') as f:
            log_content = f.read()

        lines = log_content.strip().split("\n")
        times = []

        # 記錄每行是否符合條件的詳細信息
        debug_info = []

        for i, line in enumerate(lines):
            line_info = {
                "line_num": i + 1,
                "line": line,
                "contains_prefix": prefix in line,
                "contains_suffix": suffix in line,
                "matched": False
            }

            # 檢查行是否同時包含前綴和後綴
            if prefix in line and suffix in line:
                # 找到前綴的所有位置
                prefix_positions = []
                start = 0
                while True:
                    pos = line.find(prefix, start)
                    if pos == -1:
                        break
                    prefix_positions.append(pos)
                    start = pos + 1

                # 找到後綴的所有位置
                suffix_positions = []
                start = 0
                while True:
                    pos = line.find(suffix, start)
                    if pos == -1:
                        break
                    suffix_positions.append(pos)
                    start = pos + 1

                line_info["prefix_positions"] = prefix_positions
                line_info["suffix_positions"] = suffix_positions

                # 尋找前綴後面、後綴前面的數字
                found = False
                for prefix_pos in prefix_positions:
                    prefix_end = prefix_pos + len(prefix)

                    for suffix_pos in suffix_positions:
                        # 確保前綴在後綴之前
                        if prefix_end <= suffix_pos:
                            # 提取前綴和後綴之間的文本
                            between_text = line[prefix_end:suffix_pos].strip()

                            # 直接從中間文本提取數字
                            match = re.search(r'(\d+)', between_text)

                            if match:
                                try:
                                    time_value = int(match.group(1))
                                    line_info["matched"] = True
                                    line_info["time_value"] = time_value
                                    line_info["between_text"] = between_text

                                    times.append({
                                        "file": os.path.basename(log_file),
                                        "line_num": i + 1,
                                        "time": time_value,
                                        "line": line.strip()
                                    })

                                    found = True
                                    break
                                except (ValueError, IndexError) as e:
                                    logger.warning(f"無法解析檔案 {log_file} 行 {i + 1} 中的時間值: {e}")

                        if found:
                            break

            debug_info.append(line_info)

        # 如果沒有找到匹配，記錄更詳細的調試信息
        if not times and debug_info:
            logger.debug(f"檔案 {log_file} 中找不到匹配項，調試信息:")
            for info in debug_info:
                logger.debug(f"行 {info['line_num']}: '{info['line']}'")
                logger.debug(f"  包含前綴 '{prefix}': {info['contains_prefix']}")
                logger.debug(f"  包含後綴 '{suffix}': {info['contains_suffix']}")
                if 'between_text' in info:
                    logger.debug(f"  前綴和後綴之間的文本: '{info['between_text']}'")

        return times, debug_info
    except Exception as e:
        logger.error(f"處理檔案 {log_file} 時出錯: {e}")
        return [], []


def extract_times_from_directory(log_dir, prefix, suffix):
    """從資料夾中的所有檔案提取執行時間 (不包含子資料夾，排除.exe和.bat)"""
    all_files = get_all_files(log_dir)

    if not all_files:
        logger.warning(f"在 {log_dir} 中未找到任何可分析的檔案")
        return [], {}

    logger.info(f"找到 {len(all_files)} 個檔案")
    logger.info("搜尋範圍: 僅當前資料夾，排除.exe和.bat檔案")

    all_times = []
    all_debug_info = {}
    processed_files = 0

    for log_file in all_files:
        logger.info(f"處理檔案 {processed_files + 1}/{len(all_files)}: {os.path.basename(log_file)}")
        times, debug_info = extract_times_from_file(log_file, prefix, suffix)
        all_times.extend(times)
        all_debug_info[os.path.basename(log_file)] = debug_info
        processed_files += 1

        # 每處理10個檔案顯示一次進度
        if processed_files % 10 == 0 or processed_files == len(all_files):
            logger.info(f"已處理 {processed_files}/{len(all_files)} 個檔案，找到 {len(all_times)} 個執行時間記錄")

    return all_times, all_debug_info


def calculate_statistics(times):
    """計算統計數據"""
    if not times:
        return {
            "count": 0,
            "min": 0,
            "max": 0,
            "avg": 0,
            "median": 0,
            "p90": 0,
            "p95": 0,
            "p99": 0
        }

    time_values = [item["time"] for item in times]

    return {
        "count": len(time_values),
        "min": min(time_values),
        "max": max(time_values),
        "avg": sum(time_values) / len(time_values),
        "median": np.median(time_values),
        "p90": np.percentile(time_values, 90),
        "p95": np.percentile(time_values, 95),
        "p99": np.percentile(time_values, 99)
    }


def group_by_file(times):
    """按檔案分組統計數據"""
    file_groups = {}

    for item in times:
        file_name = item["file"]
        if file_name not in file_groups:
            file_groups[file_name] = []
        file_groups[file_name].append(item["time"])

    stats = {}
    for file_name, values in file_groups.items():
        stats[file_name] = {
            "count": len(values),
            "min": min(values),
            "max": max(values),
            "avg": sum(values) / len(values),
            "median": np.median(values),
            "p90": np.percentile(values, 90),
            "p95": np.percentile(values, 95),
            "p99": np.percentile(values, 99)
        }

    return stats


def export_to_excel(times, overall_stats, file_stats, prefix, suffix, output_file):
    """將分析結果匯出到Excel檔案"""
    # 創建Excel寫入器
    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        workbook = writer.book

        # 建立格式
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })

        cell_format = workbook.add_format({
            'border': 1
        })

        number_format = workbook.add_format({
            'border': 1,
            'num_format': '0.00'
        })

        # 1. 整體統計頁
        overall_df = pd.DataFrame([{
            "找到記錄數": overall_stats["count"],
            "最小值 (ms)": overall_stats["min"],
            "最大值 (ms)": overall_stats["max"],
            "平均值 (ms)": overall_stats["avg"],
            "中位數 (ms)": overall_stats["median"],
            "90分位數 (ms)": overall_stats["p90"],
            "95分位數 (ms)": overall_stats["p95"],
            "99分位數 (ms)": overall_stats["p99"]
        }])

        overall_df.to_excel(writer, sheet_name="整體統計", index=False)
        sheet = writer.sheets["整體統計"]

        # 設置欄寬
        for i, col in enumerate(overall_df.columns):
            sheet.set_column(i, i, 15)

        # 設置格式
        for i, col in enumerate(overall_df.columns):
            sheet.write(0, i, col, header_format)
            if "ms" in col:
                sheet.write(1, i, overall_df.iloc[0, i], number_format)
            else:
                sheet.write(1, i, overall_df.iloc[0, i], cell_format)

        # 2. 檔案統計頁
        file_data = []
        for file_name, stats in file_stats.items():
            file_data.append({
                "檔案名稱": file_name,
                "找到記錄數": stats["count"],
                "最小值 (ms)": stats["min"],
                "最大值 (ms)": stats["max"],
                "平均值 (ms)": stats["avg"],
                "中位數 (ms)": stats["median"],
                "90分位數 (ms)": stats["p90"],
                "95分位數 (ms)": stats["p95"],
                "99分位數 (ms)": stats["p99"]
            })

        file_df = pd.DataFrame(file_data)
        file_df.to_excel(writer, sheet_name="檔案統計", index=False)
        sheet = writer.sheets["檔案統計"]

        # 設置欄寬
        sheet.set_column(0, 0, 40)  # 檔案名稱欄位寬一點
        for i in range(1, len(file_df.columns)):
            sheet.set_column(i, i, 15)

        # 設置格式
        for i, col in enumerate(file_df.columns):
            sheet.write(0, i, col, header_format)

        for row in range(len(file_df)):
            sheet.write(row + 1, 0, file_df.iloc[row, 0], cell_format)
            for col in range(1, len(file_df.columns)):
                if "ms" in file_df.columns[col]:
                    sheet.write(row + 1, col, file_df.iloc[row, col], number_format)
                else:
                    sheet.write(row + 1, col, file_df.iloc[row, col], cell_format)

        # 3. 不建立詳細結果頁 (已移除)
        # detail_data = []
        # for item in times:
        #     detail_data.append({
        #         "檔案名稱": item["file"],
        #         "行號": item["line_num"],
        #         "執行時間 (ms)": item["time"],
        #         "日誌行": item["line"]
        #     })
        #
        # detail_df = pd.DataFrame(detail_data)
        # detail_df.to_excel(writer, sheet_name="詳細結果", index=False)

        # 3. 創建執行時間分佈圖表
        if times:  # 只有當有數據時才創建圖表
            chart_sheet = workbook.add_worksheet("執行時間分佈圖")

            # 準備圖表數據
            chart_data = [item["time"] for item in times]
            chart_data.sort()  # 排序以便繪製

            # 寫入圖表數據
            chart_sheet.write(0, 0, "索引", header_format)
            chart_sheet.write(0, 1, "執行時間 (ms)", header_format)

            for i, value in enumerate(chart_data):
                chart_sheet.write(i + 1, 0, i + 1)
                chart_sheet.write(i + 1, 1, value)

            # 創建圖表
            chart = workbook.add_chart({'type': 'line'})

            # 配置圖表數據
            chart.add_series({
                'name': '執行時間',
                'categories': f'=執行時間分佈圖!$A$2:$A${len(chart_data) + 1}',
                'values': f'=執行時間分佈圖!$B$2:$B${len(chart_data) + 1}',
                'line': {'width': 1.5},
            })

            # 配置圖表
            chart.set_title({'name': '執行時間分佈'})
            chart.set_x_axis({'name': '樣本索引'})
            chart.set_y_axis({'name': '執行時間 (ms)'})
            chart.set_legend({'position': 'bottom'})

            # 插入圖表
            chart_sheet.insert_chart('D2', chart, {'x_scale': 1.5, 'y_scale': 1.5})

            # 設置欄寬
            chart_sheet.set_column(0, 0, 10)
            chart_sheet.set_column(1, 1, 15)

    return output_file


def main():
    """主函數"""
    args = parse_args()

    # 設置更詳細的日誌級別
    logger.setLevel(logging.DEBUG)

    # 檢查log資料夾是否存在
    if not os.path.isdir(args.log_dir):
        logger.error(f"找不到log資料夾: {args.log_dir}")
        return 1

    # 確保前綴和後綴都被指定
    if not args.prefix or not args.suffix:
        logger.error("必須同時指定前綴和後綴")
        return 1

    try:
        start_time = time.time()
        logger.info(f"正在分析log資料夾: {args.log_dir}")

        # 保留前綴和後綴中的空格
        prefix = args.prefix
        suffix = args.suffix
        logger.info(f"搜尋條件 - 前綴: '{prefix}', 後綴: '{suffix}'")
        logger.debug(f"前綴字節: {[ord(c) for c in prefix]}")
        logger.debug(f"後綴字節: {[ord(c) for c in suffix]}")

        # 提取執行時間
        logger.info("正在從所有檔案中提取執行時間...")
        times, debug_info = extract_times_from_directory(
            args.log_dir,
            prefix,
            suffix
        )

        if not times:
            logger.warning("未找到任何符合條件的執行時間")
            print("\n未找到任何符合搜尋條件的執行時間記錄。")
            print(f"請確認檔案中包含前綴 '{prefix}' 和後綴 '{suffix}'，且前綴和後綴之間有數字。")
            return 0  # 不將此視為錯誤，只是正常退出

        # 計算統計數據
        logger.info("正在計算統計數據...")
        overall_stats = calculate_statistics(times)
        file_stats = group_by_file(times)

        # 輸出檔案名 - 直接使用當前資料夾
        current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"執行時間分析結果_{current_time}.xlsx"

        # 匯出到Excel
        logger.info("正在生成Excel報告...")
        output_path = export_to_excel(
            times,
            overall_stats,
            file_stats,
            prefix,
            suffix,
            output_file
        )

        elapsed_time = time.time() - start_time
        logger.info(f"分析完成！共處理 {len(file_stats)} 個檔案，找到 {len(times)} 個執行時間記錄。")
        logger.info(f"統計結果已匯出到: {os.path.abspath(output_path)}")
        logger.info(f"處理時間: {elapsed_time:.2f} 秒")

        # 打印簡要結果
        print("\n--- 分析結果摘要 ---")
        print(f"搜尋條件 - 前綴: '{prefix}', 後綴: '{suffix}'")
        print(f"分析的檔案數: {len(file_stats)}")
        print(f"找到記錄數: {overall_stats['count']}")
        print(f"最小值: {overall_stats['min']} ms")
        print(f"最大值: {overall_stats['max']} ms")
        print(f"平均值: {overall_stats['avg']:.2f} ms")
        print(f"中位數: {overall_stats['median']:.2f} ms")
        print(f"95分位數: {overall_stats['p95']:.2f} ms")
        print(f"99分位數: {overall_stats['p99']:.2f} ms")

        return 0

    except Exception as e:
        logger.error(f"分析過程中發生錯誤: {e}", exc_info=True)
        return 1


if __name__ == "__main__":
    sys.exit(main())
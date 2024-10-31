import pandas as pd
import json
from tqdm import tqdm  # 导入 tqdm


def load_config(config_path):
    with open(config_path, 'r', encoding='utf-8') as file:
        config = json.load(file)
    return config


def extract_data(file_path, sheet_name, columns, filter_condition=None):
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # 获取实际存在的列
    existing_columns = [col for col in columns if col in df.columns]
    extracted_df = df[existing_columns]

    # 如果有筛选条件，进行行筛选
    if filter_condition:
        extracted_df = extracted_df[
            extracted_df.apply(lambda row: row.astype(str).str.contains(filter_condition).any(), axis=1)]

    return extracted_df


if __name__ == '__main__':
    # 加载配置
    config = load_config('config.json')

    # 从配置中提取变量
    file_path = config["file_path"]
    sheets = config["sheet_names"]  # 读取多个工作表名
    columns = config["columns"]
    output_file = config["output_file"]
    filter_condition = config.get("filter_condition")

    error_counts = {}  # 用于存储每个工作表中错误的数量

    with pd.ExcelWriter(output_file) as writer:
        for sheet in tqdm(sheets, desc="提取进度"):
            extracted_data = extract_data(file_path, sheet, columns, filter_condition)

            # 统计每个单元格中包含“错误”的数量
            error_count = extracted_data.apply(lambda col: col.astype(str).str.contains("错误").sum()).sum()
            error_counts[sheet] = error_count  # 记录错误数量

            if not extracted_data.empty:  # 检查提取的数据是否为空
                extracted_data.to_excel(writer, index=False, sheet_name=sheet)
                print(
                    f"工作表 '{sheet}' 的数据已提取，提取到的行数: {len(extracted_data)}，发现“错误”的数量: {error_count}")
            else:
                print(f"工作表 '{sheet}' 没有提取到任何数据，跳过保存。")

    print(f"所有数据已保存到 {output_file}")

import pandas as pd
import json


def load_config(config_path):
    """
    从 JSON 配置文件中加载配置

    参数:
    - config_path (str): 配置文件路径

    返回:
    - dict: 配置信息
    """
    with open(config_path, 'r', encoding='utf-8') as file:
        config = json.load(file)
    return config


def extract_data(file_path, sheet_name, columns, filter_condition=None):
    """
    从 Excel 文件中提取特定列的数据，并基于给定的条件进行筛选

    参数:
    - file_path (str): Excel 文件路径
    - sheet_name (str): 读取的工作表名称
    - columns (list): 需要提取的列名列表
    - filter_condition (str): 筛选条件

    返回:
    - DataFrame: 满足条件的行和列的数据
    """
    # 读取 Excel 文件
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # 提取指定列
    extracted_df = df[columns]

    # 如果有筛选条件，进行行筛选
    if filter_condition:
        # 仅保留包含筛选条件的行
        extracted_df = extracted_df[
            extracted_df.apply(lambda row: row.astype(str).str.contains(filter_condition).any(), axis=1)]

    return extracted_df


if __name__ == '__main__':
    # 加载配置
    config = load_config('config.json')

    # 从配置中提取变量
    file_path = config["file_path"]
    sheet_name = config["sheet_name"]
    columns = config["columns"]
    output_file = config["output_file"]
    output_sheet_name = config["output_sheet_name"]
    filter_condition = config.get("filter_condition")  # 筛选条件

    # 提取数据
    extracted_data = extract_data(file_path, sheet_name, columns, filter_condition)

    # 将提取的数据保存到新的 Excel 文件，并设置表单名
    with pd.ExcelWriter(output_file) as writer:
        extracted_data.to_excel(writer, index=False, sheet_name=output_sheet_name)
    print(f"Data has been extracted and saved to {output_file} with sheet name '{output_sheet_name}'")

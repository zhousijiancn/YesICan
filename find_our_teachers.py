import pandas as pd
import os

def extract_name_list(excel_file_path, sheet_name, name_column):
    """
    从Excel文件中提取指定工作表的姓名列，转换为列表
    :param excel_file_path: Excel文件的路径（相对/绝对）
    :param sheet_name: 目标工作表名称（如"固定人员名单"）
    :param name_column: 姓名列的列名（如"姓名"、"姓名/名称"等）
    :return: 姓名列表（去除空值和重复项）
    """
    # 检查文件是否存在
    if not os.path.exists(excel_file_path):
        print(f"错误：文件 {excel_file_path} 不存在，请检查路径！")
        return []

    try:
        # 读取Excel文件的指定工作表（使用openpyxl引擎处理.xlsx）
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name, engine='openpyxl')

        # 检查姓名列是否存在
        if name_column not in df.columns:
            print(f"错误：工作表 {sheet_name} 中未找到列名 {name_column}，请检查列名！")
            # 打印所有列名，方便用户核对
            print(f"当前工作表的列名有：{list(df.columns)}")
            return []

        # 提取姓名列，去除空值（NaN），转换为列表，可选去重
        name_series = df[name_column].dropna()  # 去除空值
        name_list = name_series.tolist()  # 转为列表
        # 可选：去除重复的姓名（如果有重复需要去重，注释掉则保留原顺序和重复项）
        # name_list = list(dict.fromkeys(name_list))  # 去重且保留顺序

        return name_list

    except Exception as e:
        print(f"读取文件时发生错误：{str(e)}")
        return []


def process_author_unit_column(file_path, sheet_name):
    """
    读取Excel文件的论文sheet，处理“全部作者姓名及单位”列的每一行数据
    步骤：忽略第一行 → 第二行作为列索引 → 提取目标列 → 逐行处理
    """
    # 配置文件和列信息（根据实际情况调整）
    target_cols = ["（医学部）全部第一作者姓名及单位", "（医学部）全部通讯作者姓名及单位", "全部作者姓名及单位"]  # 目标列名，需与Excel第二行的列名一致
    # 拆分目标列（明确一作、通讯、全部作者列，增强可读性）
    first_author_col = target_cols[0]
    corr_author_col = target_cols[1]
    all_author_col = target_cols[2]

    try:
        # 1. 读取Excel文件：忽略第一行，将第二行设为列索引
        df = pd.read_excel(
            file_path,
            sheet_name=sheet_name,
            skiprows=1,  # 忽略原文件的第一行
            header=0,  # 原文件的第二行作为列索引
            engine="openpyxl"  # 读取xlsx文件的引擎
        )


        # 初始化status列，默认值为1（兜底逻辑）
        df["status"] = None

        # 4. 遍历每一行处理数据
        for row_idx, row in df.iterrows():

            # 4.1 获取各列原始内容并清洗
            # 一作列内容
            first_author_content = str(row[first_author_col]).strip() if pd.notna(row[first_author_col]) else ""
            # 通讯作者列内容
            corr_author_content = str(row[corr_author_col]).strip() if pd.notna(row[corr_author_col]) else ""
            # 全部作者列内容（可选，保留用于详情展示）
            all_author_content = str(row[all_author_col]).strip() if pd.notna(row[all_author_col]) else ""

            # 4.2 核心匹配逻辑：检查一作或通讯作者列是否包含name_list中的名字
            first_has_name = any(name in first_author_content for name in names)
            corr_has_name = any(name in corr_author_content for name in names)
            all_has_name = any(name in all_author_content for name in names)

            status = None

            # 规则1：只要其一包含，status设为0
            if all_has_name:
                status = 0
            # 规则1：只要一作通讯包含，status设为1
            if first_has_name or corr_has_name:
                status = 1

            # 更新DataFrame的status列
            df.loc[row_idx, "status"] = status

        return df

    # 异常处理：捕获常见错误并给出提示
    except FileNotFoundError:
        print(f"错误：未找到文件 '{file_path}'，请检查文件路径是否正确")
        return None, []
    except Exception as e:
        print(f"处理数据时发生错误：{str(e)}")
        return None, []


# -------------------------- 主程序 --------------------------
if __name__ == "__main__":
    # 提取姓名列表
    names = extract_name_list("./副本2025年度广东省重点实验室考核评估申报书-人员和论文信息.xlsx", "固定人员清单", "姓名")

    # 处理Excel文件
    processed_result = process_author_unit_column("./2023-生工科研成果-年报-2024.4.26.xlsx", "论文" )

    # 将列表转为DataFrame
    result_df = pd.DataFrame(processed_result)
    # 调整列的顺序：确保status在第二列
    columns = result_df.columns.tolist()
    # 移除status列，插入到第二个位置
    columns.remove("status")
    columns.insert(1, "status")
    result_df = result_df[columns]

    # 使用openpyxl引擎（必选，支持UTF-8和中文），不设置index避免多余行
    result_df.to_excel(
        "./2023处理结果详情表最终版.xlsx",
        index=False,  # 不保存DataFrame的行索引
        engine="openpyxl"  # 关键引擎，确保中文正常显示
    )

    print("\n处理结果详情表已保存：处理结果详情表.xlsx（status列在第二列，支持中文显示）")


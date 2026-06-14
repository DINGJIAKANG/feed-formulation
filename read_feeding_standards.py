import os

import pandas as pd


#定义函数 read_feeding_standards，接收一个参数 file_path（Excel 文件的路径），返回整理后的字典数据
def read_feeding_standards(file_path):
    # 检查文件是否存在
    if not os.path.exists(file_path):
        print(f"错误：文件 '{file_path}' 不存在")
        return {}

    try:
        # 读取Excel文件中的所有工作表
        standards = {}
        #pd.ExcelFile(file_path)：创建一个 ExcelFile对象，用于获取 Excel 文件的所有工作表信息（如工作表名）
        excel_file = pd.ExcelFile(file_path)

        # 遍历每个工作表
        #excel_file.sheet_names：获取 Excel 文件中所有工作表的名称
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            # 将第一列设置为索引并转置
            #将 DataFrame 的第一列（df.columns[0]，通常是“营养成分”或“指标”列）设置为行索引（index）。inplace=True表示直接修改原 DataFrame，不生成新对象。
            df.set_index(df.columns[0], inplace=True)
            df = df.T

            # 转换为字典格式
            animal_standards = {}
            for stage, row in df.iterrows():
                # 处理NaN值
                stage_standards = {nutrient: value if not pd.isna(value) else None
                                   for nutrient, value in row.items()}
                animal_standards[stage] = stage_standards

            standards[sheet_name] = animal_standards

        return standards

    except Exception as e:
        print(f"读取错误: {str(e)}")
        return {}


# 测试代码
# if __name__ == "__main__":
#     # 使用您指定的文件路径
#     excel_file = r"D:\pythonfeed\feedingstandard\feedingstandard1.xlsx"
#     print(f"正在读取: {excel_file}")
#     standards = read_feeding_standards(excel_file)
#     print(standards)
    #
    # if standards:
    #     print("\n成功读取饲养标准数据:")
    #     for animal, stages in standards.items():
    #         print(f"\n动物: {animal}")
    #         for stage, nutrients in stages.items():
    #             print(f"  阶段: {stage}")
    #             for nutrient, value in nutrients.items():
    #                 print(f"    {nutrient}: {value}")
    # else:
    #     print("\n读取失败，请检查文件路径和格式")



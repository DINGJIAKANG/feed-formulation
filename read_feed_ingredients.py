import pandas as pd
import os


def read_feed_ingredients(file_path):
    """
    简化版饲料原料读取函数
    注意：第1行第2列为原料价格（元/公斤）
    新增：自动识别"精粗类型"列（若有），作为单独字段返回（不纳入营养成分）

    参数:
        file_path (str): Excel文件路径

    返回:
        dict: 包含所有饲料原料数据的字典
    """
    try:
        # 检查文件是否存在
        if not os.path.exists(file_path):
            print(f"错误：文件 '{file_path}' 不存在")
            return {}

        # 读取Excel文件
        df = pd.read_excel(file_path)

        # ── 列名标准化（兼容模板/数据库不同列名）──
        # 价格列别名
        price_aliases = ["价格", "价格(元/kg)", "价格（元/kg）", "价格(元/公斤)", "price"]
        for alias in price_aliases:
            if alias in df.columns and alias != "价格":
                df = df.rename(columns={alias: "价格"})
                break

        # 原料名称列别名
        name_aliases = ["原料名称", "饲料原料名称", "原料名", "名称", "饲料原料", "ingredient"]
        for alias in name_aliases:
            if alias in df.columns and alias != "原料名称":
                df = df.rename(columns={alias: "原料名称"})
                break

        # 精粗类型列别名
        ft_aliases = ["精粗类型", "饲料类型", "类型", "feed_type", "type"]
        for alias in ft_aliases:
            if alias in df.columns and alias != "精粗类型":
                df = df.rename(columns={alias: "精粗类型"})
                break

        # 检测特殊列（按列名，而非位置）
        has_feed_type = "精粗类型" in df.columns

        # 营养成分列 = 全列 减去 原料名称、价格、精粗类型
        special_cols = {"原料名称", "价格"}
        if has_feed_type:
            special_cols.add("精粗类型")
        nutrient_cols = [c for c in df.columns if c not in special_cols]

        # 初始化存储所有原料数据的字典
        ingredients_data = {}

        # 遍历每一行数据
        for index, row in df.iterrows():
            # 使用列名访问（避免iloc的FutureWarning）
            ingredient_name = row["原料名称"]
            price = row["价格"]

            # 获取精粗类型（若有）
            feed_type = row["精粗类型"] if has_feed_type else None

            # 获取营养成分数据（排除特殊列）
            nutrients = {}
            for col in nutrient_cols:
                val = row[col]
                nutrients[col] = val if not pd.isna(val) else None

            # 组装返回字典
            item = {
                '价格': price,
                '营养成分': nutrients,
            }
            if feed_type is not None:
                ft_str = str(feed_type).strip()
                # 只保留有效值（"精料"/"粗料"/"其他"，空值不记录）
                if ft_str and ft_str.lower() not in ("nan", "none", ""):
                    item['精粗类型'] = ft_str

            # 添加到原料数据字典
            ingredients_data[ingredient_name] = item

        return ingredients_data

    except Exception as e:
        print(f"读取错误: {str(e)}")
        return {}


def display_ingredients(ingredients):
    """
    显示饲料原料数据（包括价格）
    """
    if not ingredients:
        print("没有可显示的饲料原料数据")
        return

    print("\n" + "=" * 60)
    print("饲料原料数据（含价格信息）")
    print("=" * 60)

    for ingredient, data in ingredients.items():
        price = data['价格']
        nutrients = data['营养成分']
        feed_type = data.get('精粗类型', '未分类')

        print(f"\n原料: {ingredient}  [{feed_type}]")
        print(f"  价格: {price} 元/公斤" if not pd.isna(price) else "  价格: 未指定")

        for nutrient, value in nutrients.items():
            display_value = "未指定" if pd.isna(value) else value
            print(f"  ├─ {nutrient}: {display_value}")

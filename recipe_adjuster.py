"""
配方微调模块
功能：在 LP 最优解基础上，允许用户手动微调各原料配比，
     实时计算营养指标变化，对比原始配方与微调配方的差异。

设计原则：
  - 独立模块，对现有代码零侵入
  - 输入：result dict（build_and_solve 返回值）+ ingredients_df（原料 DataFrame）
  - 输出：对比 DataFrame（原始 vs 微调 + 标准）
"""
from __future__ import annotations

import numpy as np
import pandas as pd


def calc_nutrients(ingredients_pct: dict, ingredients_df: pd.DataFrame,
                   nutrient_map: dict | None = None) -> dict:
    """
    给定原料配比(%)和原料DataFrame，计算各营养指标含量。

    参数:
        ingredients_pct: {原料名: 百分比}，百分比总和应为100
        ingredients_df: 原料营养数据（index=原料名，columns=营养指标）
        nutrient_map: {标准列名: 原料列名} 映射，若为None则使用所有数值列

    返回:
        {营养指标: 含量}（含量已归一化到单位比例，即百分比÷100后加权平均）
    """
    if not ingredients_pct or ingredients_df is None:
        return {}

    total = sum(float(v) for v in ingredients_pct.values())
    if total <= 0:
        return {}

    nutrients = {}

    if nutrient_map:
        # 只计算 nutrient_map 中映射的列
        for std_col, ing_col in nutrient_map.items():
            if ing_col not in ingredients_df.columns:
                continue
            vals = pd.to_numeric(ingredients_df[ing_col], errors="coerce").fillna(0)
            weighted_sum = 0.0
            for name, pct in ingredients_pct.items():
                real_name = name
                if name.startswith("[上传] "):
                    real_name = name[5:]
                elif name.startswith("[默认] "):
                    real_name = name[5:]
                if real_name in vals.index:
                    weighted_sum += float(vals.loc[real_name]) * float(pct) / total
            nutrients[std_col] = round(weighted_sum, 4)
    else:
        # 计算所有数值列
        for col in ingredients_df.columns:
            if col in ("价格", "编号", "描述", "排序号", "精粗类型"):
                continue
            vals = pd.to_numeric(ingredients_df[col], errors="coerce")
            if vals.notna().sum() == 0:
                continue
            vals = vals.fillna(0)
            weighted_sum = 0.0
            for name, pct in ingredients_pct.items():
                real_name = name
                if name.startswith("[上传] "):
                    real_name = name[5:]
                elif name.startswith("[默认] "):
                    real_name = name[5:]
                if real_name in vals.index:
                    weighted_sum += float(vals.loc[real_name]) * float(pct) / total
            nutrients[col] = round(weighted_sum, 4)

    return nutrients


def calc_cost(ingredients_pct: dict, ingredients_df: pd.DataFrame) -> float:
    """
    计算配方成本（元/kg）。
    公式：Σ(配比% × 单价) / 100
    """
    if not ingredients_pct or ingredients_df is None:
        return 0.0

    prices = pd.to_numeric(ingredients_df["价格"], errors="coerce").fillna(0)
    total = sum(float(v) for v in ingredients_pct.values())
    if total <= 0:
        return 0.0

    cost = 0.0
    for name, pct in ingredients_pct.items():
        real_name = name
        if name.startswith("[上传] "):
            real_name = name[5:]
        elif name.startswith("[默认] "):
            real_name = name[5:]
        if real_name in prices.index:
            cost += float(prices.loc[real_name]) * float(pct) / total

    return round(cost, 4)


def calc_concentrate_forage(ingredients_pct: dict, ingredients_df: pd.DataFrame) -> dict:
    """
    计算微调配方的精粗比。
    返回: {"concentrate_pct": xx.x, "forage_pct": xx.x} 或空dict
    优先使用「精粗类型」列，若不存在则用关键词后备分类。
    """
    total = sum(float(v) for v in ingredients_pct.values())
    if total <= 0:
        return {}

    # 粗料关键词（用于后备分类）
    _FORAGE_KEYWORDS = ("草", "秸秆", "苜蓿", "青贮", "干草", "牧草", "稻草",
                        "麦秸", "玉米秸", "豆秸", "花生秧", "羊草",
                        "树叶", "甘蔗渣", "甜菜粕", "纤维")

    concentrate = 0.0
    forage = 0.0
    has_type_col = "精粗类型" in ingredients_df.columns

    for name, pct in ingredients_pct.items():
        real_name = name
        if name.startswith("[上传] "):
            real_name = name[5:]
        elif name.startswith("[默认] "):
            real_name = name[5:]

        if has_type_col and real_name in ingredients_df.index:
            ft = str(ingredients_df.at[real_name, "精粗类型"]).strip()
            if ft == "精料":
                concentrate += float(pct)
            elif ft == "粗料":
                forage += float(pct)
        elif not has_type_col:
            # 后备：根据原料名称关键词判断
            _is_forage = any(kw in real_name for kw in _FORAGE_KEYWORDS)
            if _is_forage:
                forage += float(pct)
            else:
                concentrate += float(pct)

    return {
        "concentrate_pct": round(concentrate / total * 100, 2),
        "forage_pct": round(forage / total * 100, 2),
    }


def calc_ca_p_ratio(ingredients_pct: dict, ingredients_df: pd.DataFrame) -> float | None:
    """计算钙磷比。"""
    if "钙%" not in ingredients_df.columns or "总磷%" not in ingredients_df.columns:
        return None

    total = sum(float(v) for v in ingredients_pct.values())
    if total <= 0:
        return None

    ca_vals = pd.to_numeric(ingredients_df["钙%"], errors="coerce").fillna(0)
    p_vals = pd.to_numeric(ingredients_df["总磷%"], errors="coerce").fillna(0)

    ca_sum = 0.0
    p_sum = 0.0
    for name, pct in ingredients_pct.items():
        real_name = name
        if name.startswith("[上传] "):
            real_name = name[5:]
        elif name.startswith("[默认] "):
            real_name = name[5:]
        if real_name in ca_vals.index:
            ca_sum += float(ca_vals.loc[real_name]) * float(pct) / total
        if real_name in p_vals.index:
            p_sum += float(p_vals.loc[real_name]) * float(pct) / total

    if p_sum > 1e-10:
        return round(ca_sum / p_sum, 2)
    return None


def build_comparison_df(original_result: dict, adjusted_pct: dict,
                        ingredients_df: pd.DataFrame,
                        nutrient_map: dict | None = None) -> pd.DataFrame:
    """
    构建原始配方 vs 微调配方的营养对比表。

    参数:
        original_result: build_and_solve() 返回的 result dict
        adjusted_pct: 微调后的 {原料名: 百分比}
        ingredients_df: 原料 DataFrame
        nutrient_map: 营养指标映射（与求解时一致）

    返回:
        DataFrame 列: [营养指标, 标准需要量, 原始配方, 微调配方, 变化量, 达成]
    """
    requirements = original_result.get("requirements", {})
    original_nutrients = original_result.get("nutrients", {})
    adjusted_nutrients = calc_nutrients(adjusted_pct, ingredients_df, nutrient_map)

    rows = []
    for std_col, req_val in requirements.items():
        if req_val is None:
            continue
        try:
            req_float = float(req_val)
        except (ValueError, TypeError):
            continue

        orig_val = original_nutrients.get(std_col, 0.0)
        adj_val = adjusted_nutrients.get(std_col, 0.0)
        diff = adj_val - orig_val  # 微调 - 原始
        meets = "✅" if round(adj_val - req_float, 4) >= 0 else "❌"

        rows.append({
            "营养指标": std_col,
            "标准需要量": req_float,
            "原始配方": orig_val,
            "微调配方": adj_val,
            "变化量": round(diff, 4),
            "达成": meets,
        })

    return pd.DataFrame(rows)


# ── 表格编辑式微调 ──────────────────────────────────────────


def _strip_tag(name: str) -> str:
    """去掉 [上传] / [默认] 标签前缀"""
    if name.startswith("[上传] "):
        return name[5:]
    elif name.startswith("[默认] "):
        return name[5:]
    return name


def build_adjustment_editable_df(
    original_pct: dict,
    current_adj_pct: dict,
    ingredients_df: pd.DataFrame,
    requirements: dict | None = None,
    nutrient_map: dict | None = None,
    max_nutrient_cols: int = 8,
) -> tuple[pd.DataFrame, list[str]]:
    """
    构建可编辑的微调数据表（用于 st.data_editor）。

    返回:
        (editable_df, nutrient_col_names):
          editable_df 包含列：原料, 原配比%, 微调配比%, 营养1贡献, 营养2贡献, ...
          nutrient_col_names: 营养贡献列名列表（用于后续汇总计算）
    """
    # 确定要展示的营养指标（优先从requirements取，最多max_nutrient_cols个）
    _reqs = requirements or {}
    _nmap = nutrient_map or {}

    # 营养列优先级：能量 > 粗蛋白 > 钙磷 > 赖氨酸蛋氨酸 > 其他
    _priority_keys = []
    _energy_keywords = ["消化能", "代谢能", "净能"]
    _key_nutrients = ["粗蛋白%", "钙%", "总磷%", "有效磷%", "赖氨酸%", "蛋氨酸%", "色氨酸%", "苏氨酸%"]

    # 先找能量
    for k in sorted(_reqs.keys()):
        if any(e_kw in k for e_kw in _energy_keywords):
            _priority_keys.append(k)

    # 再找关键营养素
    for kn in _key_nutrients:
        for k in sorted(_reqs.keys()):
            if k == kn or (kn.rstrip("%") in k and "%" in k):
                if k not in _priority_keys:
                    _priority_keys.append(k)

    # 其余营养素
    for k in sorted(_reqs.keys()):
        if k not in _priority_keys:
            _priority_keys.append(k)

    _show_nutrients = _priority_keys[:max_nutrient_cols]

    # 构建每行的数据
    rows = []
    _sorted_ingredients = sorted(
        original_pct.keys(), key=lambda k: -float(original_pct[k])
    )

    for _ing_name in _sorted_ingredients:
        _orig_val = float(original_pct.get(_ing_name, 0))
        _adj_val = float(current_adj_pct.get(_ing_name, _orig_val))

        row_data = {
            "原料": _ing_name,
            "原配比%": round(_orig_val, 3),
            "微调配比%": round(_adj_val, 3),
        }

        # 计算该原料对各营养的贡献值 = 配比% × 营养含量 / 100
        _real_name = _strip_tag(_ing_name)
        for _ncol in _show_nutrients:
            _ing_col = _nmap.get(_ncol, _ncol)
            if _ing_col in ingredients_df.columns and _real_name in ingredients_df.index:
                _nval = pd.to_numeric(
                    ingredients_df.at[_real_name, _ing_col], errors="coerce"
                )
                _contrib = _adj_val * (_nval if pd.notna(_nval) else 0) / 100.0
                row_data[_ncol] = round(_contrib, 4)
            else:
                row_data[_ncol] = 0.0

        rows.append(row_data)

    df = pd.DataFrame(rows)
    return df, _show_nutrients


def calc_summary_rows(
    adj_pct: dict,
    ingredients_df: pd.DataFrame,
    nutrient_names: list[str],
    requirements: dict | None = None,
    nutrient_map: dict | None = None,
) -> list[dict]:
    """计算底部汇总行（合计、标准、差）。返回行字典列表。"""
    total = sum(float(v) for v in adj_pct.values())
    if total <= 0:
        total = 100.0  # 防止除零

    _nmap = nutrient_map or {}
    _reqs = requirements or {}

    # 合计行
    _adj_nuts = calc_nutrients(adj_pct, ingredients_df, _nmap)
    total_row = {"合计": ""}
    for nn in nutrient_names:
        total_row[nn] = round(_adj_nuts.get(nn, 0), 4)

    # 标准行
    std_row = {"标准": ""}
    for nn in nutrient_names:
        val = _reqs.get(nn)
        try:
            std_row[nn] = round(float(val), 4) if val is not None else ""
        except (ValueError, TypeError):
            std_row[nn] = ""

    # 与标准的差
    diff_row = {"与标准的差": ""}
    for nn in nutrient_names:
        t = total_row.get(nn, 0)
        s = std_row.get(nn, "")
        try:
            diff_row[nn] = round(t - float(s), 4) if s != "" else ""
        except (ValueError, TypeError):
            diff_row[nn] = ""

    return [total_row, std_row, diff_row]

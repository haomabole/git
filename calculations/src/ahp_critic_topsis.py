"""AHP + CRITIC + TOPSIS evaluation workflow.

Configure the constants in the first section, then run:

    python calculations/src/ahp_critic_topsis.py
"""

from pathlib import Path

import numpy as np
import pandas as pd


# ============================================================
# 1. User configuration
# ============================================================

INPUT_FILE = "input.xlsx"
SHEET_NAME = "Sheet1"
SAMPLE_COL = "试样编号"
OUTPUT_FILE = "output.xlsx"

# Weight combination method:
# - "additive": w = alpha * w_AHP + (1 - alpha) * w_CRITIC
# - "multiplicative": w = (w_AHP)^alpha * (w_CRITIC)^(1 - alpha)
COMBINE_METHOD = "multiplicative"

# Larger alpha means the final weight leans more toward subjective AHP weights.
ALPHA = 0.50
EPS = 1e-12

CRITIC_NORMALIZE_METHOD = "minmax"
TOPSIS_NORMALIZE_METHOD = "minmax"

# direction:
# - "max": larger is better
# - "min": smaller is better
INDICATORS = [
    {"name": "抗拉强度", "col": "抗拉强度/MPa", "direction": "max"},
    {"name": "硬度", "col": "硬度/HV", "direction": "max"},
    {"name": "延伸率", "col": "延伸率/%", "direction": "max"},
]

# Subjective AHP percentages. Set to None to use AHP_MATRIX instead.
AHP_BY_PERCENT = {
    "抗拉强度": 60,
    "硬度": 20,
    "延伸率": 20,
}

AHP_MATRIX = [
    [1, 3, 3],
    [1 / 3, 1, 1],
    [1 / 3, 1, 1],
]

# Optional direct subjective/final weights.
SUBJECTIVE_WEIGHTS = None
MANUAL_FINAL_WEIGHTS = None

# Optional final weight caps, for example: {"延伸率": 0.20}
WEIGHT_CAPS = {}


# ============================================================
# 2. General utilities
# ============================================================

RI_TABLE = {
    1: 0.00,
    2: 0.00,
    3: 0.58,
    4: 0.90,
    5: 1.12,
    6: 1.24,
    7: 1.32,
    8: 1.41,
    9: 1.45,
    10: 1.49,
    11: 1.51,
    12: 1.48,
    13: 1.56,
    14: 1.57,
    15: 1.59,
}


def normalize_weights(weights):
    weights = pd.Series(weights, dtype=float)
    if (weights < 0).any():
        raise ValueError("权重中不能出现负数。")

    total = weights.sum()
    if total <= 0:
        raise ValueError("权重之和必须大于 0。")

    return weights / total


def build_ahp_matrix_from_percent(percent_dict, items):
    weights = pd.Series(percent_dict, dtype=float).reindex(items)

    if weights.isna().any():
        missing = weights[weights.isna()].index.tolist()
        raise ValueError(f"AHP_BY_PERCENT 中缺少指标：{missing}")

    if (weights <= 0).any():
        raise ValueError("AHP_BY_PERCENT 中所有权重必须大于 0。")

    weights = normalize_weights(weights)

    n = len(items)
    matrix = np.ones((n, n), dtype=float)
    for i in range(n):
        for j in range(n):
            matrix[i, j] = weights.iloc[i] / weights.iloc[j]

    return matrix.tolist()


def calculate_ahp_weights(matrix, items):
    matrix_array = np.array(matrix, dtype=float)
    n = len(items)

    if matrix_array.shape != (n, n):
        raise ValueError(f"AHP 矩阵维度应为 {n}x{n}，当前为 {matrix_array.shape}。")

    if not np.allclose(matrix_array * matrix_array.T, np.ones_like(matrix_array), atol=1e-6):
        raise ValueError("AHP 判断矩阵不是互反矩阵，请检查 a_ij 和 a_ji 是否互为倒数。")

    eigvals, eigvecs = np.linalg.eig(matrix_array)
    max_index = np.argmax(eigvals.real)
    lambda_max = eigvals[max_index].real
    principal_vec = eigvecs[:, max_index].real

    weights = np.abs(principal_vec)
    weights = weights / weights.sum()
    ahp_weights = pd.Series(weights, index=items, name="AHP主观权重")

    if n <= 2:
        ci, ri, cr, passed = 0.0, 0.0, 0.0, True
    else:
        ci = (lambda_max - n) / (n - 1)
        ri = RI_TABLE.get(n, 1.59)
        cr = ci / ri if ri != 0 else 0.0
        passed = cr < 0.10

    consistency = pd.DataFrame(
        {
            "指标数量n": [n],
            "lambda_max": [lambda_max],
            "CI": [ci],
            "RI": [ri],
            "CR": [cr],
            "是否通过CR<0.10": [passed],
        }
    )

    matrix_df = pd.DataFrame(matrix_array, index=items, columns=items)
    return ahp_weights, consistency, matrix_df


# ============================================================
# 3. Indicator normalization
# ============================================================

def normalize_indicator_matrix(df, indicators, method="minmax"):
    score_df = pd.DataFrame(index=df.index)

    for item in indicators:
        name = item["name"]
        col = item["col"]
        direction = item["direction"]

        if col not in df.columns:
            raise KeyError(f"Excel 中找不到列：{col}")

        values = pd.to_numeric(df[col], errors="coerce")
        if values.isna().any():
            raise ValueError(f"列 {col} 中存在空值或非数值，请先清理数据。")

        if method == "minmax":
            xmin = values.min()
            xmax = values.max()

            if np.isclose(xmax, xmin):
                normalized = pd.Series(np.ones(len(values)), index=values.index)
            elif direction == "max":
                normalized = (values - xmin) / (xmax - xmin)
            elif direction == "min":
                normalized = (xmax - values) / (xmax - xmin)
            else:
                raise ValueError(f"{name} 的 direction 必须是 'max' 或 'min'。")

        elif method == "desirability":
            r = item.get("r", 1)

            if direction == "max":
                low = item.get("low", values.min())
                target = item.get("target", values.max())
                if np.isclose(target, low):
                    raise ValueError(f"{name} 的 target 和 low 不能相等。")
                normalized = ((values - low) / (target - low)).clip(0, 1) ** r

            elif direction == "min":
                target = item.get("target", values.min())
                upper = item.get("upper", values.max())
                if np.isclose(upper, target):
                    raise ValueError(f"{name} 的 upper 和 target 不能相等。")
                normalized = ((upper - values) / (upper - target)).clip(0, 1) ** r

            else:
                raise ValueError(f"{name} 的 direction 必须是 'max' 或 'min'。")

        else:
            raise ValueError("method 必须是 'minmax' 或 'desirability'。")

        score_df[name] = normalized

    return score_df


# ============================================================
# 4. CRITIC objective weights
# ============================================================

def calculate_critic_weights(score_df):
    matrix = score_df.copy().astype(float)
    std = matrix.std(axis=0, ddof=0)
    corr = matrix.corr().fillna(0.0)

    critic_info = {}
    for col in matrix.columns:
        conflict = sum(1 - corr.loc[col, other] for other in matrix.columns if other != col)
        information = std[col] * conflict
        critic_info[col] = {
            "标准差": std[col],
            "冲突性": conflict,
            "信息量Cj": information,
        }

    info_df = pd.DataFrame(critic_info).T
    info_sum = info_df["信息量Cj"].sum()

    if np.isclose(info_sum, 0):
        weights = pd.Series(
            np.ones(len(score_df.columns)) / len(score_df.columns),
            index=score_df.columns,
            name="CRITIC客观权重",
        )
    else:
        weights = info_df["信息量Cj"] / info_sum
        weights.name = "CRITIC客观权重"

    return weights, info_df, corr


# ============================================================
# 5. Combined weights
# ============================================================

def combine_weights(subjective_weights, objective_weights, method="multiplicative", alpha=0.5, eps=1e-12):
    subjective = normalize_weights(subjective_weights)
    objective = normalize_weights(objective_weights)

    if method == "additive":
        final = alpha * subjective + (1 - alpha) * objective
    elif method == "multiplicative":
        final = (subjective + eps) ** alpha * (objective + eps) ** (1 - alpha)
    else:
        raise ValueError("COMBINE_METHOD 必须是 'additive' 或 'multiplicative'。")

    final = normalize_weights(final)
    final.name = f"最终组合权重_{method}"
    return final


def apply_weight_caps(weights, caps):
    if not caps:
        return normalize_weights(weights)

    original = normalize_weights(weights)
    capped = original.copy()
    fixed = pd.Series(False, index=capped.index)
    caps_series = pd.Series(caps, dtype=float).reindex(capped.index)

    for _ in range(len(capped) + 1):
        over = (caps_series.notna()) & (capped > caps_series) & (~fixed)
        if not over.any():
            break

        fixed[over] = True
        capped[over] = caps_series[over]

        fixed_sum = capped[fixed].sum()
        remaining_total = 1.0 - fixed_sum
        if remaining_total < -1e-12:
            raise ValueError("权重上限设置过低，导致固定权重之和超过 1。")

        free = ~fixed
        if free.any():
            base = original[free]
            capped[free] = remaining_total * base / base.sum()

    return normalize_weights(capped)


# ============================================================
# 6. Weighted TOPSIS
# ============================================================

def weighted_topsis(score_df, weights, use_fixed_ideal=True):
    items = score_df.columns.tolist()
    weights = normalize_weights(weights).reindex(items)

    if use_fixed_ideal:
        ideal_best = pd.Series(1.0, index=items, name="正理想解")
        ideal_worst = pd.Series(0.0, index=items, name="负理想解")
    else:
        ideal_best = score_df.max(axis=0)
        ideal_best.name = "正理想解"
        ideal_worst = score_df.min(axis=0)
        ideal_worst.name = "负理想解"

    d_plus = np.sqrt(((score_df - ideal_best) ** 2 * weights).sum(axis=1))
    d_minus = np.sqrt(((score_df - ideal_worst) ** 2 * weights).sum(axis=1))
    closeness = d_minus / (d_plus + d_minus)
    score_100 = 100 * closeness

    result = pd.DataFrame(
        {
            "D_plus_到正理想解距离": d_plus,
            "D_minus_到负理想解距离": d_minus,
            "TOPSIS贴近度": closeness,
            "TOPSIS评分_100分": score_100,
        },
        index=score_df.index,
    )

    ideal_df = pd.concat([ideal_best, ideal_worst], axis=1)

    contribution_plus = (score_df - ideal_best) ** 2 * weights
    contribution_plus.columns = [f"Dplus贡献_{col}" for col in contribution_plus.columns]

    contribution_minus = (score_df - ideal_worst) ** 2 * weights
    contribution_minus.columns = [f"Dminus贡献_{col}" for col in contribution_minus.columns]

    return result, ideal_df, contribution_plus, contribution_minus


# ============================================================
# 7. Main workflow
# ============================================================

def run_evaluation():
    items = [item["name"] for item in INDICATORS]
    cols = [item["col"] for item in INDICATORS]

    if not Path(INPUT_FILE).exists():
        raise FileNotFoundError(f"找不到输入文件：{INPUT_FILE}")

    df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)
    if SAMPLE_COL not in df.columns:
        raise KeyError(f"Excel 中找不到试样编号列：{SAMPLE_COL}")

    critic_score_df = normalize_indicator_matrix(
        df,
        INDICATORS,
        method=CRITIC_NORMALIZE_METHOD,
    )
    topsis_score_df = normalize_indicator_matrix(
        df,
        INDICATORS,
        method=TOPSIS_NORMALIZE_METHOD,
    )

    if SUBJECTIVE_WEIGHTS is not None:
        subject_weights = pd.Series(SUBJECTIVE_WEIGHTS, dtype=float).reindex(items)
        subject_weights = normalize_weights(subject_weights)
        subject_weights.name = "AHP/主观权重"
        consistency_df = pd.DataFrame({"说明": ["使用直接主观权重，未进行 AHP 一致性检验。"]})
        ahp_matrix_df = pd.DataFrame()
    else:
        ahp_matrix = build_ahp_matrix_from_percent(AHP_BY_PERCENT, items) if AHP_BY_PERCENT is not None else AHP_MATRIX
        subject_weights, consistency_df, ahp_matrix_df = calculate_ahp_weights(ahp_matrix, items)
        subject_weights.name = "AHP/主观权重"

    critic_weights, critic_detail, corr_df = calculate_critic_weights(critic_score_df)
    critic_weights = critic_weights.reindex(items)
    critic_weights.name = "CRITIC客观权重"

    if MANUAL_FINAL_WEIGHTS is not None:
        final_weights = pd.Series(MANUAL_FINAL_WEIGHTS, dtype=float).reindex(items)
        final_weights = normalize_weights(final_weights)
        final_weights.name = "最终组合权重_手动"
    else:
        final_weights = combine_weights(
            subject_weights,
            critic_weights,
            method=COMBINE_METHOD,
            alpha=ALPHA,
            eps=EPS,
        )
        final_weights = apply_weight_caps(final_weights, WEIGHT_CAPS)
        final_weights.name = f"最终组合权重_{COMBINE_METHOD}"

    topsis_result, ideal_df, contribution_plus, contribution_minus = weighted_topsis(
        topsis_score_df,
        final_weights,
        use_fixed_ideal=True,
    )

    raw_result = df[[SAMPLE_COL] + cols].copy()
    topsis_norm_out = topsis_score_df.copy()
    topsis_norm_out.columns = [f"topsis_d_{col}" for col in topsis_norm_out.columns]

    result_df = pd.concat(
        [raw_result, topsis_norm_out, topsis_result, contribution_plus, contribution_minus],
        axis=1,
    )
    result_df["排名"] = result_df["TOPSIS贴近度"].rank(ascending=False, method="min").astype(int)
    result_df = result_df.sort_values("TOPSIS贴近度", ascending=False)

    weights_df = pd.concat([subject_weights, critic_weights, final_weights], axis=1)
    weights_df["最终权重_%"] = weights_df[final_weights.name] * 100

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        result_df.to_excel(writer, sheet_name="topsis_result", index=False)
        weights_df.to_excel(writer, sheet_name="weights")
        consistency_df.to_excel(writer, sheet_name="ahp_consistency", index=False)
        ahp_matrix_df.to_excel(writer, sheet_name="ahp_matrix")
        critic_detail.to_excel(writer, sheet_name="critic_detail")
        corr_df.to_excel(writer, sheet_name="correlation_matrix")
        critic_score_df.to_excel(writer, sheet_name="critic_normalized")
        topsis_score_df.to_excel(writer, sheet_name="topsis_normalized")
        ideal_df.to_excel(writer, sheet_name="topsis_ideal_solutions")
        contribution_plus.to_excel(writer, sheet_name="Dplus_contribution")
        contribution_minus.to_excel(writer, sheet_name="Dminus_contribution")

    print("计算完成。")
    print(f"输出文件：{OUTPUT_FILE}")
    print("\n最终权重：")
    print(weights_df)
    print("\nAHP一致性检验：")
    print(consistency_df)
    print("\nTOPSIS评分前5名：")
    print(result_df[[SAMPLE_COL, "TOPSIS贴近度", "TOPSIS评分_100分", "排名"]].head())

    return result_df, weights_df


if __name__ == "__main__":
    run_evaluation()


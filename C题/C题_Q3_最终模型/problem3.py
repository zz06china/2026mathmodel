"""
C题 问题3：基于光伏预测误差的滚动优化 + 调整购电计价机制。

核心机制（相对问题2新增）：
  - 每天 0:00 / 6:00 / 12:00 / 18:00 各获得一次未来24小时整点光伏预报（附件3）。
  - 0:00 制定计划购电量 B_t；6/12/18 可据此调整，得到调整购电量 A_t。
  - 调整购电量与计划购电量的差量按特殊价结算：
        Δ_up   = max(0, A_t - B_t)   -> 超出部分 1.5 倍电价
        Δ_down = max(0, B_t - A_t)   -> 违约部分 50%  电价（即净省 0.5p）
  - 总费用 = 计划购电费 + 紧急购电费 + 调整购电量相关费用
        C = Σ p_t B_t + Σ 1.5 p_t Δ_up_t - Σ 0.5 p_t Δ_down_t + Σ 5 p_t R_t
          = Σ p_t A_t + Σ 0.5 p_t (Δ_up_t + Δ_down_t) + Σ 5 p_t R_t   （等价）

模型结构（顺序MPC + 因果执行）：
  1. 0:00：用0:00预报解日前LP -> B_t（锁定，作为调整基准）。
  2. 6:00：用6:00预报+实际SOC，对剩余时段重优化（含调整成本）-> A_t[6:00-24:00]，
     执行 6:00-12:00。
  3. 12:00、18:00 同理。
  4. 逐时段因果执行：实际PV/负载观测下优先充放电，缺口才紧急购电。
"""

from __future__ import annotations

from copy import copy
from pathlib import Path
from shutil import copyfile

import numpy as np
import openpyxl
import pulp

BASE = Path(__file__).resolve().parent.parent  # C题/
ATTACHMENT1 = BASE / "附件" / "附件1.xlsx"
ATTACHMENT2 = BASE / "附件" / "附件2.xlsx"
ATTACHMENT3 = BASE / "附件" / "附件3.xlsx"
TEMPLATE = BASE / "附件" / "附件5" / "result3.xlsx"
OUTPUT = Path(__file__).resolve().parent / "result3.xlsx"
SUMMARY = Path(__file__).resolve().parent / "q3_metrics.txt"

DT = 1.0 / 6.0
N = 144
ETA_C = 0.9
ETA_D = 0.9
E_MIN = 1200.0
E_MAX = 10800.0
E_INITIAL = 6000.0
Q_MAX = 5000.0 * DT
REPORT_START = 31
ALPHA_GRID = np.linspace(0.0, 1.0, 11)
SOLVER = pulp.PULP_CBC_CMD(msg=False)
TOL = 1e-5
EXPORT_TOL = 5e-5
LEX_TOL = 1e-3  # 词典序第二阶段成本容差（元）
ISSUE_HOURS = [0, 6, 12, 18]


# ---------------------------------------------------------------------------
# 时间轴与数据加载
# ---------------------------------------------------------------------------
def rotate_to_natural_day(values):
    values = list(values)
    if len(values) != N:
        raise ValueError(f"Expected {N} intervals, got {len(values)}")
    return np.asarray([values[-1], *values[:-1]], dtype=float)


def source_order(values):
    values = list(values)
    return [*values[1:], values[0]]


def load_inputs():
    wb1 = openpyxl.load_workbook(ATTACHMENT1, data_only=True, read_only=True)
    rows = list(wb1["Sheet1"].iter_rows(min_row=2, values_only=True))
    wb1.close()
    prices = rotate_to_natural_day(float(row[1]) for row in rows)
    prior_load = rotate_to_natural_day(float(row[2]) * DT for row in rows)
    prior_pv = rotate_to_natural_day(float(row[3]) * DT for row in rows)

    wb2 = openpyxl.load_workbook(ATTACHMENT2, data_only=True, read_only=True)
    load_sheet, pv_sheet = wb2.worksheets[:2]

    def read_sheet(sheet):
        data, dates = [], []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[0] is None:
                continue
            dates.append(row[0])
            data.append(rotate_to_natural_day(float(v) * DT for v in row[1 : N + 1]))
        return dates, np.asarray(data)

    dates, loads = read_sheet(load_sheet)
    pv_dates, pvs = read_sheet(pv_sheet)
    wb2.close()
    return prices, prior_load, prior_pv, dates, loads, pvs


def load_forecasts():
    """附件3：返回 shape (365, 4, 24)，[day][issue][预报1..24小时]."""
    wb = openpyxl.load_workbook(ATTACHMENT3, data_only=True, read_only=True)
    ws = wb["Sheet1"]
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    wb.close()
    assert len(rows) == 365 * 4, f"附件3行数={len(rows)}"
    forecasts = np.zeros((365, 4, 24))

    def issue_index(label):
        hour = int(str(label).strip().split(":")[0])
        return {0: 0, 6: 1, 12: 2, 18: 3}[hour]

    day = -1
    for r in rows:
        if r[0] is not None and str(r[0]).strip() != "":
            day += 1
        idx = issue_index(r[1])
        forecasts[day, idx] = np.array([float(v) for v in r[2:26]])
    return forecasts


# ---------------------------------------------------------------------------
# 预测：负载（自适应滞后融合，沿用第二问）与光伏（附件3插值）
# ---------------------------------------------------------------------------
def select_alpha(data, day, prior):
    max_validation_day = day - 1
    start = max(7, max_validation_day - 55)
    candidates = []
    for alpha in ALPHA_GRID:
        errors = []
        for k in range(start, max_validation_day + 1):
            pred = alpha * data[k - 1] + (1.0 - alpha) * data[k - 7]
            errors.append(np.mean(np.abs(data[k] - pred)))
        candidates.append(np.mean(errors) if errors else np.inf)
    if not np.isfinite(min(candidates)):
        return 0.0
    return float(ALPHA_GRID[int(np.argmin(candidates))])


def load_forecast_10min(loads, day, prior_load):
    """日负荷10分钟预测（只在0:00做一次，日内不更新）."""
    if day < 7:
        return prior_load.copy()
    alpha = select_alpha(loads, day, prior_load)
    return np.maximum(alpha * loads[day - 1] + (1.0 - alpha) * loads[day - 7], 0.0)


def residual_margin(residual_history, window=28, q=0.80):
    """净负荷预测残差的 q 分位修正（报童安全库存）.

    短缺成本 5p、超购（弃电/违约退款）成本 ~0，一阶条件 F(q)=1-1/5=0.8。
    """
    if not residual_history:
        return np.zeros(N)
    sample = np.asarray(residual_history[-window:])
    return np.quantile(sample, q, axis=0)


def hourly_forecast_to_10min(hourly, issue_hour):
    """把24个整点预报值插值到144个10分钟时段。

    hourly[k] = 光伏在 (issue_hour + k + 1) 点的预报值，k=0..23。
    当前小时（issue_hour）用 hourly[0] 作为近似；更早的时段置 0。
    """
    marker = np.zeros(25)  # marker[h] = 小时 h 的光伏能量（kWh/10min = kW * DT）
    for h in range(1, 25):
        off = h - issue_hour
        if 1 <= off <= 24:
            marker[h] = hourly[off - 1] * DT
    if issue_hour > 0:
        marker[issue_hour] = hourly[0] * DT
    minutes = (np.arange(N) * 10 + 5) / 60.0  # 每时段中点在“小时”单位的坐标
    return np.interp(minutes, np.arange(25), marker)


# ---------------------------------------------------------------------------
# LP 求解
# ---------------------------------------------------------------------------
def _finish_two_stage(model, cost, smoothing_vars):
    """第一阶段最小化 cost，第二阶段在容差内最小化吞吐量/弃电以去退化."""
    model.setObjective(cost)
    status = model.solve(SOLVER)
    if pulp.LpStatus[status] != "Optimal":
        raise RuntimeError(f"第一阶段求解失败: {pulp.LpStatus[status]}")
    optimum = pulp.value(cost)
    model += cost <= optimum + LEX_TOL
    model.setObjective(pulp.lpSum(smoothing_vars))
    status = model.solve(SOLVER)
    if pulp.LpStatus[status] != "Optimal":
        raise RuntimeError(f"第二阶段求解失败: {pulp.LpStatus[status]}")


def solve_dayahead(net_targets, prices, e_start, terminal=E_INITIAL):
    """0:00 日前计划：min Σ p B，得到计划购电量 B_t（144时段）."""
    model = pulp.LpProblem("Q3_dayahead", pulp.LpMinimize)
    buy = pulp.LpVariable.dicts("buy", range(N), lowBound=0)
    charge = pulp.LpVariable.dicts("charge", range(N), 0, Q_MAX)
    discharge = pulp.LpVariable.dicts("discharge", range(N), 0, Q_MAX)
    surplus = pulp.LpVariable.dicts("surplus", range(N), lowBound=0)
    energy = pulp.LpVariable.dicts("energy", range(N + 1), E_MIN, E_MAX)

    model += energy[0] == e_start
    model += energy[N] == terminal
    for t in range(N):
        model += buy[t] + discharge[t] == float(net_targets[t]) + charge[t] + surplus[t]
        model += energy[t + 1] == energy[t] + ETA_C * charge[t] - discharge[t] / ETA_D

    cost = pulp.lpSum(prices[t] * buy[t] for t in range(N))
    _finish_two_stage(model, cost, [charge[t] + discharge[t] + surplus[t] for t in range(N)])
    return {
        "buy": np.array([pulp.value(buy[t]) for t in range(N)]),
        "charge": np.array([pulp.value(charge[t]) for t in range(N)]),
        "discharge": np.array([pulp.value(discharge[t]) for t in range(N)]),
        "surplus": np.array([pulp.value(surplus[t]) for t in range(N)]),
        "energy": np.array([pulp.value(energy[t]) for t in range(N + 1)]),
    }


def solve_adjust(B_full, net_input, prices, e_start, start_slot, terminal=E_INITIAL):
    """在 start_slot 时刻重优化剩余时段，含相对计划 B 的调整成本。

    返回长度 (N - start_slot) 的调整购电量 A_t。
    目标：min Σ [p A + 0.5 p (Δ_up + Δ_down)]，约束 A - Δ_up + Δ_down = B。
    由于上调价 1.5p、下调价 -0.5p（相对计划），该目标对 A 是凸的分段线性函数，
    纯 LP 即可（无需 0-1 变量），最优解自动满足 Δ_up·Δ_down=0。
    """
    S = start_slot
    M = N - S
    model = pulp.LpProblem(f"Q3_adjust_{S}", pulp.LpMinimize)
    buy = pulp.LpVariable.dicts("a", range(M), lowBound=0)
    charge = pulp.LpVariable.dicts("charge", range(M), 0, Q_MAX)
    discharge = pulp.LpVariable.dicts("discharge", range(M), 0, Q_MAX)
    surplus = pulp.LpVariable.dicts("surplus", range(M), lowBound=0)
    dup = pulp.LpVariable.dicts("dup", range(M), lowBound=0)
    ddown = pulp.LpVariable.dicts("ddown", range(M), lowBound=0)
    energy = pulp.LpVariable.dicts("energy", range(M + 1), E_MIN, E_MAX)

    model += energy[0] == e_start
    model += energy[M] == terminal
    for i in range(M):
        t = S + i
        model += buy[i] - dup[i] + ddown[i] == float(B_full[t])
        model += (
            buy[i] + discharge[i]
            == float(net_input[t]) + charge[i] + surplus[i]
        )
        model += energy[i + 1] == energy[i] + ETA_C * charge[i] - discharge[i] / ETA_D

    cost = pulp.lpSum(
        prices[S + i] * (buy[i] + 0.5 * (dup[i] + ddown[i])) for i in range(M)
    )
    _finish_two_stage(
        model, cost, [charge[i] + discharge[i] + surplus[i] for i in range(M)]
    )
    return np.array([pulp.value(buy[i]) for i in range(M)])


# ---------------------------------------------------------------------------
# 因果执行
# ---------------------------------------------------------------------------
def execute_range(buy, load, pv, energy, start, stop, charge, discharge, emergency, surplus):
    """在 [start, stop) 时段按实际观测因果平衡，缺口才形成紧急购电."""
    for i in range(start, stop):
        net_deficit = load[i] - pv[i] - buy[i]
        if net_deficit >= 0.0:
            discharge[i] = min(net_deficit, Q_MAX, max((energy - E_MIN) * ETA_D, 0.0))
            emergency[i] = net_deficit - discharge[i]
            charge[i] = 0.0
            surplus[i] = 0.0
        else:
            available = -net_deficit
            charge[i] = min(available, Q_MAX, max((E_MAX - energy) / ETA_C, 0.0))
            surplus[i] = available - charge[i]
            discharge[i] = 0.0
            emergency[i] = 0.0
        energy += ETA_C * charge[i] - discharge[i] / ETA_D
    return energy


def validate_day(charge, discharge, energy, emergency, surplus, buy, load, pv):
    balance = buy + pv + discharge + emergency - load - charge - surplus
    storage = energy[1:] - energy[:-1] - ETA_C * charge + discharge / ETA_D
    simultaneous = np.flatnonzero((charge > TOL) & (discharge > TOL))
    return {
        "balance": float(np.max(np.abs(balance))),
        "storage": float(np.max(np.abs(storage))),
        "min_energy": float(np.min(energy)),
        "max_energy": float(np.max(energy)),
        "simultaneous": int(len(simultaneous)),
    }


# ---------------------------------------------------------------------------
# 输出
# ---------------------------------------------------------------------------
def interval_label(start, stop):
    def hhmm(minutes):
        minutes %= 1440
        return f"{minutes // 60:02d}:{minutes % 60:02d}"

    start_min = start * 10
    stop_min = stop * 10
    suffix = "+1" if stop_min >= 1440 else ""
    return f"{hhmm(start_min)}-{hhmm(stop_min)}{suffix}"


def copy_row_style(sheet, source_row, target_row, max_col):
    from copy import copy as _copy

    sheet.row_dimensions[target_row].height = sheet.row_dimensions[source_row].height
    for col in range(1, max_col + 1):
        src = sheet.cell(source_row, col)
        dst = sheet.cell(target_row, col)
        if src.has_style:
            dst._style = _copy(src._style)
        dst.number_format = src.number_format
        dst.alignment = _copy(src.alignment)
        dst.font = _copy(src.font)
        dst.fill = _copy(src.fill)
        dst.border = _copy(src.border)


def write_output(dates, prices, results):
    copyfile(TEMPLATE, OUTPUT)
    wb = openpyxl.load_workbook(OUTPUT)

    # 计划购电量：B_t；最后两列 全天购电量 / 全天购电费（计划费）
    ws = wb["计划购电量"]
    for i, r in enumerate(results):
        row = i + 2
        for t, v in enumerate(source_order(r["buy"]), start=2):
            ws.cell(row, t, round(float(v), 4))
        ws.cell(row, 146, round(float(np.sum(r["buy"])), 4))
        ws.cell(row, 147, round(float(np.dot(prices, r["buy"])), 4))

    # 调整购电量：A_t；最后两列 全天调整购电量 / 调整购电量相关费用
    ws = wb["调整购电量"]
    for i, r in enumerate(results):
        row = i + 2
        for t, v in enumerate(source_order(r["adjusted"]), start=2):
            ws.cell(row, t, round(float(v), 4))
        ws.cell(row, 146, round(float(np.sum(r["adjusted"])), 4))
        ws.cell(row, 147, round(float(r["adjust_cost"]), 4))

    # 充放电量 + 储电量
    ws = wb["充放电量"]
    style_rows = [[copy(ws.cell(r, c)._style) for c in range(1, 7)] for r in range(2, 8)]
    ws.delete_rows(2, ws.max_row)
    labels = ["0:00-4:00", "4:00-8:00", "8:00-12:00", "12:00-16:00", "16:00-20:00", "20:00-24:00"]
    row = 2
    for date, r in zip(dates, results):
        for block in range(6):
            for col in range(1, 7):
                ws.cell(row, col)._style = copy(style_rows[block][col - 1])
            start, stop = block * 24, (block + 1) * 24
            ws.cell(row, 1, date if block == 0 else None)
            ws.cell(row, 2, labels[block])
            ws.cell(row, 3, round(float(np.sum(r["charge"][start:stop])), 4))
            ws.cell(row, 4, round(float(np.sum(r["discharge"][start:stop])), 4))
            if block == 0:
                ws.cell(row, 5, "0:00")
                ws.cell(row, 6, round(float(r["energy"][0]), 4))
            elif block == 1:
                ws.cell(row, 5, "24:00")
                ws.cell(row, 6, round(float(r["energy"][-1]), 4))
            row += 1

    # 紧急购电量
    ws = wb["紧急购电量"]
    style_rows = [[copy(ws.cell(r, c)._style) for c in range(1, 4)] for r in range(2, 5)]
    ws.delete_rows(2, ws.max_row)
    row = 2
    for date, r in zip(dates, results):
        emergency = r["emergency"]
        starts, t = [], 0
        while t < N:
            if emergency[t] < EXPORT_TOL:
                t += 1
                continue
            start = t
            while t < N and emergency[t] >= EXPORT_TOL:
                t += 1
            starts.append((start, t))
        for ei, (start, stop) in enumerate(starts):
            style = style_rows[min(ei, len(style_rows) - 1)]
            for col in range(1, 4):
                ws.cell(row, col)._style = copy(style[col - 1])
            ws.cell(row, 1, date if ei == 0 else None)
            ws.cell(row, 2, interval_label(start, stop))
            ws.cell(row, 3, round(float(np.sum(emergency[start:stop])), 4))
            row += 1

    wb.save(OUTPUT)


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def run_model(adjust_issues):
    """按给定调整时刻集合运行滚动优化，返回 (report 结果列表, 指标 dict)。

    adjust_issues 为 [6,12,18] 的子集；空集即"无日内调整"（纯日前）对照。
    每个时段 t 的购电量在"承诺时刻"（≤ t//6 的最近一次预报时刻）锁定，
    安全裕量残差也按该承诺时刻的预报（0-6h 前）计算，而非 0:00 整日预报。
    """
    prices, prior_load, prior_pv, dates, loads, pvs = load_inputs()
    forecasts = load_forecasts()
    adjust_issues = sorted(adjust_issues)
    boundaries = [h * 6 for h in adjust_issues]  # 如 [36, 72, 108]
    first_boundary = boundaries[0] if boundaries else N

    all_results = []
    residuals = []
    e_start = E_INITIAL
    max_checks = {"balance": 0.0, "storage": 0.0, "min_energy": E_MAX, "max_energy": E_MIN}

    for day in range(len(dates)):
        load0 = load_forecast_10min(loads, day, prior_load)
        pv = {0: hourly_forecast_to_10min(forecasts[day, 0], 0)}
        for h in adjust_issues:
            pv[h] = hourly_forecast_to_10min(forecasts[day, h // 6], h)
        margin = residual_margin(residuals)

        # 0:00 计划
        B = solve_dayahead(load0 - pv[0] + margin, prices, e_start)["buy"]

        charge = np.zeros(N)
        discharge = np.zeros(N)
        emergency = np.zeros(N)
        surplus = np.zeros(N)
        energy = np.zeros(N + 1)
        energy[0] = e_start
        A = B.copy()  # 最终调整购电量，先默认等于计划

        # 逐段执行 + 在承诺边界处调整
        e = e_start
        prev = 0
        for b in boundaries:
            e = execute_range(A, loads[day], pvs[day], e, prev, b, charge, discharge, emergency, surplus)
            h = b // 6
            net = np.concatenate(
                [loads[day][:b] - pvs[day][:b], load0[b:] - pv[h][b:] + margin[b:]]
            )
            A[b:144] = solve_adjust(B, net, prices, e, b)
            prev = b
        e = execute_range(A, loads[day], pvs[day], e, prev, N, charge, discharge, emergency, surplus)

        energy = _build_energy_trajectory(e_start, charge, discharge)

        # 差量与费用（调整费仅发生在首个承诺边界之后的时段）
        delta = A - B
        dup = np.maximum(delta, 0.0)
        ddown = np.maximum(-delta, 0.0)
        plan_cost = float(np.dot(prices, B))
        adjust_cost = float(np.dot(prices[first_boundary:], 1.5 * dup[first_boundary:] - 0.5 * ddown[first_boundary:]))
        emergency_cost = float(np.dot(5.0 * prices, emergency))

        checks = validate_day(charge, discharge, energy, emergency, surplus, A, loads[day], pvs[day])
        max_checks["balance"] = max(max_checks["balance"], checks["balance"])
        max_checks["storage"] = max(max_checks["storage"], checks["storage"])
        max_checks["min_energy"] = min(max_checks["min_energy"], checks["min_energy"])
        max_checks["max_energy"] = max(max_checks["max_energy"], checks["max_energy"])

        all_results.append(
            {
                "buy": B,
                "adjusted": A,
                "charge": charge,
                "discharge": discharge,
                "emergency": emergency,
                "surplus": surplus,
                "energy": energy,
                "plan_cost": plan_cost,
                "adjust_cost": adjust_cost,
                "emergency_cost": emergency_cost,
                "dup": dup,
                "ddown": ddown,
            }
        )

        # 承诺时刻残差：每段用"锁定该段购电量的预报"（0-6h 前），而非 0:00 整日预报
        residual = np.zeros(N)
        commit_points = [0] + boundaries  # [0, 36, 72, 108]
        for i, b in enumerate(commit_points):
            b_next = commit_points[i + 1] if i + 1 < len(commit_points) else N
            ch = b // 6
            residual[b:b_next] = (loads[day] - pvs[day])[b:b_next] - (load0 - pv[ch])[b:b_next]
        residuals.append(residual)
        e_start = float(energy[-1])

    report = all_results[REPORT_START:]
    report_dates = dates[REPORT_START:]

    plan_kwh = sum(np.sum(r["buy"]) for r in report)
    adjusted_kwh = sum(np.sum(r["adjusted"]) for r in report)
    dup_kwh = sum(np.sum(r["dup"]) for r in report)
    ddown_kwh = sum(np.sum(r["ddown"]) for r in report)
    emergency_kwh = sum(np.sum(r["emergency"]) for r in report)
    plan_cost = sum(r["plan_cost"] for r in report)
    adjust_cost = sum(r["adjust_cost"] for r in report)
    emergency_cost = sum(r["emergency_cost"] for r in report)
    total_cost = plan_cost + adjust_cost + emergency_cost
    emergency_days = sum(np.sum(r["emergency"]) > 0.01 for r in report)
    adjust_days = sum(np.sum(r["dup"] + r["ddown"]) > 0.01 for r in report)

    metrics = {
        "report_days": len(report),
        "planned_purchase": plan_kwh,
        "adjusted_purchase": adjusted_kwh,
        "up_adjust": dup_kwh,
        "down_adjust": ddown_kwh,
        "emergency_purchase": emergency_kwh,
        "planned_cost": plan_cost,
        "adjust_cost": adjust_cost,
        "emergency_cost": emergency_cost,
        "total_cost": total_cost,
        "emergency_days": emergency_days,
        "adjust_days": adjust_days,
        "ending_energy": float(report[-1]["energy"][-1]),
        "max_balance_residual": max_checks["balance"],
        "max_storage_residual": max_checks["storage"],
        "global_min_energy": max_checks["min_energy"],
        "global_max_energy": max_checks["max_energy"],
        "simultaneous": sum(int(np.sum((r["charge"] > TOL) & (r["discharge"] > TOL))) for r in report),
    }
    return report, report_dates, metrics


def main():
    prices, _, _, _, _, _ = load_inputs()
    report, report_dates, metrics = run_model([6, 12, 18])
    write_output(report_dates, prices, report)

    lines = [
        "Q3 MODEL METRICS (MPC + 调整购电计价)",
        f"report_days={metrics['report_days']}",
        f"planned_purchase={metrics['planned_purchase']:.4f} kWh",
        f"adjusted_purchase={metrics['adjusted_purchase']:.4f} kWh",
        f"up_adjust={metrics['up_adjust']:.4f} kWh",
        f"down_adjust={metrics['down_adjust']:.4f} kWh",
        f"emergency_purchase={metrics['emergency_purchase']:.4f} kWh",
        f"planned_cost={metrics['planned_cost']:.4f} yuan",
        f"adjust_cost={metrics['adjust_cost']:.4f} yuan",
        f"emergency_cost={metrics['emergency_cost']:.4f} yuan",
        f"total_cost={metrics['total_cost']:.4f} yuan",
        f"emergency_days={metrics['emergency_days']}/{metrics['report_days']}",
        f"adjust_days={metrics['adjust_days']}/{metrics['report_days']}",
        f"ending_energy={metrics['ending_energy']:.4f} kWh",
        f"max_balance_residual={metrics['max_balance_residual']:.8g}",
        f"max_storage_residual={metrics['max_storage_residual']:.8g}",
        f"global_min_energy={metrics['global_min_energy']:.4f} kWh",
        f"global_max_energy={metrics['global_max_energy']:.4f} kWh",
        f"simultaneous={metrics['simultaneous']}",
        f"output={OUTPUT}",
    ]
    SUMMARY.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("\n".join(lines))


def marginal_analysis():
    """对比不同日内调整时刻集合的总费用，回答"是否需要引入其他时刻预报"."""
    print(f"{'adjust_issues':<18}{'planned_cost':>14}{'adjust_cost':>13}{'emergency':>12}{'total_cost':>13}")
    for issues in [[], [6], [6, 12], [6, 12, 18]]:
        _, _, m = run_model(issues)
        print(
            f"{str(issues):<18}{m['planned_cost']:>14.0f}{m['adjust_cost']:>13.0f}"
            f"{m['emergency_cost']:>12.0f}{m['total_cost']:>13.0f}"
        )


def _build_energy_trajectory(e_start, charge, discharge):
    traj = np.zeros(N + 1)
    traj[0] = e_start
    for t in range(N):
        traj[t + 1] = traj[t] + ETA_C * charge[t] - discharge[t] / ETA_D
    return traj


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1 and sys.argv[1] == "marginal":
        marginal_analysis()
    else:
        main()

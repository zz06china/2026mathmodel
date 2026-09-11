# -*- coding: utf-8 -*-
"""
问题1：微网计划购电策略 —— 线性规划 (LP)
求解器：PuLP + CBC

模型：
  min  Σ p_t · b_t
  s.t. b_t + G_t + d_t = L_t + c_t          (能量平衡)
       E_t = E_{t-1} + η·c_t - d_t/η        (储能递推, η=0.9)
       1200 ≤ E_t ≤ 10800                    (SOC上下限)
       0 ≤ c_t, d_t ≤ 5000·Δt                (充放电功率上限)
       b_t ≥ 0
       E_0 = E_144 = 6000                    (期初=期末)
"""
import openpyxl
import pulp

# ---------- 1. 读数据 ----------
wb = openpyxl.load_workbook("附件/附件1.xlsx", data_only=True)
ws = wb["Sheet1"]
rows = list(ws.iter_rows(values_only=True))
header, data = rows[0], rows[1:]

DT = 1 / 6.0          # 每个时段 10 分钟 = 1/6 小时
T = len(data)         # 144 个时段

# 时间标签 -> 从0:00起的分钟数 (处理最后一行的 '0:00+1')
def time_to_minutes(v):
    if isinstance(v, str):               # '0:00+1' 表示当天 24:00
        if v.strip() == "0:00+1":
            return 24 * 60
        hh, mm = v.split(":")
        return int(hh) * 60 + int(mm)
    return v.hour * 60 + v.minute         # datetime.time

times = [time_to_minutes(r[0]) for r in data]
prices = [float(r[1]) for r in data]      # 元/kWh
L = [float(r[2]) * DT for r in data]      # 负载 kW -> kWh
G = [float(r[3]) * DT for r in data]      # 光伏 kW -> kWh

# ---------- 2. 储能参数 ----------
E0 = 6000.0        # 0:00 初始电量 (kWh)
E_min = 1200.0     # 运行下限
E_max = 10800.0    # 运行上限 (注意不是12000)
C_max = 5000.0 * DT  # 单时段最大充/放电能量 (kWh)
eta = 0.9          # 充放电效率

# ---------- 3. 建模 ----------
prob = pulp.LpProblem("microgrid_p1", pulp.LpMinimize)

b = [pulp.LpVariable(f"b_{t}", lowBound=0) for t in range(T)]   # 购电量
c = [pulp.LpVariable(f"c_{t}", lowBound=0) for t in range(T)]   # 充电量
d = [pulp.LpVariable(f"d_{t}", lowBound=0) for t in range(T)]   # 放电量
E = [pulp.LpVariable(f"E_{t}") for t in range(T)]               # 时段末储能电量

# 目标：全天购电费最小
prob += pulp.lpSum(prices[t] * b[t] for t in range(T))

# 约束
for t in range(T):
    # ① 能量平衡
    prob += b[t] + G[t] + d[t] == L[t] + c[t]
    # ② 储能递推
    E_prev = E0 if t == 0 else E[t - 1]
    prob += E[t] == E_prev + eta * c[t] - d[t] / eta
    # ③ SOC 上下限
    prob += E[t] >= E_min
    prob += E[t] <= E_max
    # ④ 充放电功率上限
    prob += c[t] <= C_max
    prob += d[t] <= C_max
# ⑤ 期末储能 = 期初
prob += E[T - 1] == E0

# ---------- 4. 求解 ----------
prob.solve(pulp.PULP_CBC_CMD(msg=0))
print("求解状态:", pulp.LpStatus[prob.status])
print("全天购电费(元):", round(pulp.value(prob.objective), 2))
print("全天购电量(kWh):", round(sum(pulp.value(b[t]) for t in range(T)), 2))

# ---------- 5. 结果输出 ----------
b_val = [pulp.value(b[t]) for t in range(T)]
c_val = [pulp.value(c[t]) for t in range(T)]
d_val = [pulp.value(d[t]) for t in range(T)]
E_val = [pulp.value(E[t]) for t in range(T)]

# 表1：指定10分钟时段的购电量
print("\n【表1】指定时段购电量:")
slots_10min = ["10:00-10:10", "12:00-12:10", "14:00-14:10",
               "16:00-16:10", "18:00-18:10", "20:00-20:10"]
for label in slots_10min:
    hh, mm = label.split("-")[0].split(":")
    idx = (int(hh) * 60 + int(mm)) // 10 - 1   # 该10分钟时段对应的索引
    print(f"  {label}: {b_val[idx]:.2f} kWh")

# 表2：指定4小时时段的充/放电量
print("\n【表2】指定时段充/放电量:")
blocks = [("0:00-4:00", 0), ("4:00-8:00", 4), ("8:00-12:00", 8),
          ("12:00-16:00", 12), ("16:00-20:00", 16), ("20:00-24:00", 20)]
for label, start_h in blocks:
    i0, i1 = start_h * 6, (start_h + 4) * 6   # 每小时6个时段
    chg = sum(c_val[i] for i in range(i0, i1))
    dis = sum(d_val[i] for i in range(i0, i1))
    print(f"  {label}: 充电 {chg:.2f} kWh, 放电 {dis:.2f} kWh")

print(f"\n0:00 储电量: {E0:.2f} kWh")
print(f"24:00 储电量: {E_val[-1]:.2f} kWh")

# 校验约束是否满足
max_soc = max(E_val); min_soc = min(E_val)
print(f"\n[校验] SOC 范围: {min_soc:.2f} ~ {max_soc:.2f} (应落在 1200~10800)")
print(f"[校验] 全天净充电(含效率): {sum(eta*c_val[t] - d_val[t]/eta for t in range(T)):.2f} (应为0, 期初=期末)")

# ---------- 6. 写入 result1.xlsx (按附件5模板) ----------
TEMPLATE = "附件/附件5/result1.xlsx"
OUTPUT = "result1.xlsx"

wb_out = openpyxl.load_workbook(TEMPLATE)

# 计划购电量工作表：第2列(购电量) 按顺序填入 144 个时段
ws1 = wb_out["计划购电量"]
for t in range(T):
    ws1.cell(row=t + 2, column=2, value=round(b_val[t], 4))

# 充放电量工作表：6个4小时时段的充电量/放电量 + 0:00/24:00储电量
ws2 = wb_out["充放电量"]
for k, start_h in enumerate([0, 4, 8, 12, 16, 20]):
    i0, i1 = start_h * 6, (start_h + 4) * 6   # 每小时6个时段
    chg = round(sum(c_val[i] for i in range(i0, i1)), 4)
    dis = round(sum(d_val[i] for i in range(i0, i1)), 4)
    ws2.cell(row=k + 2, column=2, value=chg)   # 充电量
    ws2.cell(row=k + 2, column=3, value=dis)   # 放电量
ws2.cell(row=2, column=5, value=round(E0, 4))          # 0:00 储电量
ws2.cell(row=3, column=5, value=round(E_val[-1], 4))   # 24:00 储电量

wb_out.save(OUTPUT)
print(f"\n已保存结果到 {OUTPUT}")

from pathlib import Path
import sys

import matplotlib.pyplot as plt
import numpy as np


HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import q4_final  # noqa: E402


FIGURE_DIR = HERE / "figures"
DATES = ("2025-03-20", "2025-06-21", "2025-09-23", "2025-12-21")
BLUE = "#4472C4"
ORANGE = "#ED7D31"
GREEN = "#70AD47"
RED = "#C64E4E"
GRAY = "#666666"


def setup():
    plt.rcParams.update(
        {
            "font.family": "sans-serif",
            "font.sans-serif": ["Microsoft YaHei", "SimHei", "Arial Unicode MS"],
            "axes.unicode_minus": False,
            "axes.linewidth": 0.8,
            "svg.fonttype": "none",
            "figure.facecolor": "white",
            "axes.facecolor": "white",
        }
    )


def price_profiles():
    _, _, _, dates, _, _ = q4_final.q2_final.load_inputs()
    prices = q4_final.load_dynamic_prices(dates)
    date_index = {str(value)[:10]: index for index, value in enumerate(dates)}
    hours = np.arange(q4_final.N) / 6
    colors = (BLUE, ORANGE, GREEN, RED)

    fig, ax = plt.subplots(figsize=(9.2, 4.0))
    for label, color in zip(DATES, colors):
        ax.plot(hours, prices[date_index[label]], color=color, linewidth=1.35, label=label)
    ax.set_xlim(0, 24)
    ax.set_xticks(np.arange(0, 25, 4))
    ax.set_xlabel("时刻/h")
    ax.set_ylabel("电价/(元·kWh$^{-1}$)")
    ax.grid(axis="y", color="#D9D9D9", linestyle="--", linewidth=0.6)
    ax.spines[["top", "right"]].set_visible(False)
    ax.legend(ncol=4, frameon=False, loc="upper center", bbox_to_anchor=(0.5, 1.13))
    fig.tight_layout()
    fig.savefig(FIGURE_DIR / "q4_price_profiles.svg", format="svg", bbox_inches="tight")
    plt.close(fig)


def optimization_and_response():
    models = ("问题4-2", "问题4-3")
    fixed_cost = np.array((1489.6533, 1444.8943))
    optimized_cost = np.array((1474.1729, 1433.5349))
    charge_price = np.array((0.5017, 0.5043))
    discharge_price = np.array((1.1365, 1.1152))

    fig, axes = plt.subplots(1, 2, figsize=(9.2, 3.8))
    x = np.arange(len(models))
    width = 0.32
    axes[0].bar(x - width / 2, fixed_cost, width, color=GRAY, label="固定电价策略动态结算")
    axes[0].bar(x + width / 2, optimized_cost, width, color=BLUE, label="动态电价优化")
    axes[0].set_xticks(x, models)
    axes[0].set_ylabel("总费用/万元")
    axes[0].set_ylim(1400, 1510)
    axes[0].grid(axis="y", color="#D9D9D9", linestyle="--", linewidth=0.6)
    axes[0].spines[["top", "right"]].set_visible(False)
    axes[0].legend(frameon=False, fontsize=8, loc="upper right")
    for bars in axes[0].containers:
        axes[0].bar_label(bars, fmt="%.1f", fontsize=8, padding=2)

    axes[1].bar(x - width / 2, charge_price, width, color=GREEN, label="加权平均充电电价")
    axes[1].bar(x + width / 2, discharge_price, width, color=ORANGE, label="加权平均放电电价")
    axes[1].set_xticks(x, models)
    axes[1].set_ylabel("电价/(元·kWh$^{-1}$)")
    axes[1].set_ylim(0, 1.3)
    axes[1].grid(axis="y", color="#D9D9D9", linestyle="--", linewidth=0.6)
    axes[1].spines[["top", "right"]].set_visible(False)
    axes[1].legend(frameon=False, fontsize=8, loc="upper left")
    for bars in axes[1].containers:
        axes[1].bar_label(bars, fmt="%.2f", fontsize=8, padding=2)

    fig.tight_layout(w_pad=2.0)
    fig.savefig(
        FIGURE_DIR / "q4_optimization_response.svg",
        format="svg",
        bbox_inches="tight",
    )
    plt.close(fig)


def main():
    FIGURE_DIR.mkdir(parents=True, exist_ok=True)
    setup()
    price_profiles()
    optimization_and_response()
    print(FIGURE_DIR / "q4_price_profiles.svg")
    print(FIGURE_DIR / "q4_optimization_response.svg")


if __name__ == "__main__":
    main()

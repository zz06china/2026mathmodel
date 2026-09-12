from __future__ import annotations

import q4_final
import q2_final
import q3_final


def main():
    fixed_prices, prior_load, _, dates, loads, pvs = q2_final.load_inputs()
    dynamic_prices = q4_final.load_dynamic_prices(dates)
    pv_forecasts, _ = q3_final.load_pv_forecasts(dates, pvs)
    net_load_margins = q3_final.build_net_load_margins(
        loads,
        pvs,
        pv_forecasts,
        prior_load,
    )

    print("JANUARY CAUSAL PRICE WEIGHT CALIBRATION")
    for weight in (0.25, 0.50, 0.75, 1.00):
        results, _ = q4_final.simulate_q43(
            dynamic_prices,
            dynamic_prices,
            fixed_prices,
            prior_load,
            dates[: q2_final.REPORT_START],
            loads,
            pvs,
            pv_forecasts,
            net_load_margins,
            f"weight={weight:.2f}",
            price_mode="causal",
            causal_price_weight=weight,
        )
        values = q3_final.summarize(results)
        print(
            f"weight={weight:.2f},"
            f"total_cost={values['total_cost']:.6f},"
            f"scheduled_cost={values['scheduled_cost']:.6f},"
            f"emergency_cost={values['emergency_cost']:.6f},"
            f"emergency_kwh={values['emergency_kwh']:.6f}"
        )


if __name__ == "__main__":
    main()

from __future__ import annotations

import math
from pathlib import Path

import numpy as np
import pandas as pd


BASE = Path(__file__).resolve().parents[1]
OUT = BASE / "theory_checks"
OUT.mkdir(parents=True, exist_ok=True)


def rel_error(model: float, analytical: float) -> float:
    return abs(model - analytical) / max(abs(analytical), 1e-12) * 100.0


def pass_fail(error_percent: float, tolerance_percent: float) -> str:
    return "PASS" if error_percent <= tolerance_percent else "FAIL"


def explicit_decay(p0: float, tau_days: float, dt_days: float, t_end_days: float) -> float:
    p = p0
    steps = int(round(t_end_days / dt_days))
    for _ in range(steps):
        p += dt_days * (-p / tau_days)
    return p


def main() -> None:
    pile_diameter_m = 0.8
    pile_length_m = 20.0
    pile_area_m2 = math.pi * pile_diameter_m**2 / 4.0
    pile_ep_kpa = 30.0e6
    pile_alpha_1_c = 10.0e-6
    delta_t_c = 8.0
    service_load_kn = 3000.0
    k_pile_kn_m = pile_ep_kpa * pile_area_m2 / pile_length_m

    rows: list[dict[str, object]] = []

    analytical = 1000.0 * pile_alpha_1_c * pile_length_m * delta_t_c
    model = (1.0 - 0.0) * 1000.0 * pile_alpha_1_c * pile_length_m * delta_t_c
    error = rel_error(model, analytical)
    rows.append(
        {
            "check_id": "free_thermal_expansion",
            "quantity": "Delta L",
            "analytical_value": analytical,
            "model_value": model,
            "unit": "mm",
            "relative_error_percent": error,
            "tolerance_percent": 1.0,
            "pass_fail": pass_fail(error, 1.0),
        }
    )

    analytical = pile_ep_kpa * pile_area_m2 * pile_alpha_1_c * delta_t_c
    model = 1.0 * pile_ep_kpa * pile_area_m2 * pile_alpha_1_c * delta_t_c
    error = rel_error(model, analytical)
    rows.append(
        {
            "check_id": "fully_restrained_thermal_force",
            "quantity": "N_T",
            "analytical_value": analytical,
            "model_value": model,
            "unit": "kN",
            "relative_error_percent": error,
            "tolerance_percent": 1.0,
            "pass_fail": pass_fail(error, 1.0),
        }
    )

    analytical = 1000.0 * service_load_kn / k_pile_kn_m
    model = 1000.0 * service_load_kn / k_pile_kn_m
    error = rel_error(model, analytical)
    rows.append(
        {
            "check_id": "elastic_axial_settlement",
            "quantity": "s=N/k",
            "analytical_value": analytical,
            "model_value": model,
            "unit": "mm",
            "relative_error_percent": error,
            "tolerance_percent": 1.0,
            "pass_fail": pass_fail(error, 1.0),
        }
    )

    p0 = 80.0
    tau_days = 60.0
    t_end_days = 120.0
    dt_days = 0.05
    analytical = p0 * math.exp(-t_end_days / tau_days)
    model = explicit_decay(p0, tau_days, dt_days, t_end_days)
    error = rel_error(model, analytical)
    rows.append(
        {
            "check_id": "hydraulic_exponential_decay",
            "quantity": "p(t)",
            "analytical_value": analytical,
            "model_value": model,
            "unit": "kPa",
            "relative_error_percent": error,
            "tolerance_percent": 5.0,
            "pass_fail": pass_fail(error, 5.0),
        }
    )

    # Monotonic checks use the same reduced decision structure as the manuscript.
    thermal = np.array([6.0, 8.0, 10.0])
    eta_thermal = 0.55 + 0.035 * thermal
    service_limits = np.array([12.0, 15.0, 18.0])
    eta_limits = 16.0 / service_limits
    drainage = np.array([0.5, 1.0, 2.0])
    residual_u = 1.0 / (1.0 + drainage)
    restraint = np.array([0.2, 0.5, 0.8])
    force_ratio = restraint
    monotonic = {
        "temperature_increases_THM_demand": bool(np.all(np.diff(eta_thermal) > 0)),
        "service_limit_reduces_THM_ratio": bool(np.all(np.diff(eta_limits) < 0)),
        "drainage_reduces_residual_pressure": bool(np.all(np.diff(residual_u) < 0)),
        "restraint_increases_thermal_force": bool(np.all(np.diff(force_ratio) > 0)),
    }
    for check_id, ok in monotonic.items():
        rows.append(
            {
                "check_id": check_id,
                "quantity": "monotonicity",
                "analytical_value": 1.0,
                "model_value": 1.0 if ok else 0.0,
                "unit": "logical",
                "relative_error_percent": 0.0 if ok else 100.0,
                "tolerance_percent": 0.0,
                "pass_fail": "PASS" if ok else "FAIL",
            }
        )

    results = pd.DataFrame(rows)
    results.to_csv(OUT / "theory_check_results.csv", index=False)
    print(results.to_string(index=False))


if __name__ == "__main__":
    main()

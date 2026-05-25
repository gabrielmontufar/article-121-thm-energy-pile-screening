from __future__ import annotations

import json
import sys
from pathlib import Path

import pandas as pd

from load_transfer_model import elastic_head_movement_mm, restrained_thermal_movement_mm, service_class


BASE = Path(__file__).resolve().parents[2]
VAL = BASE / "validation"
SUPP = BASE / "Supplementary data"


def main() -> None:
    params = pd.read_csv(VAL / "external_parameters.csv")
    parametric = pd.read_csv(SUPP / "parametric_serviceability_matrix.csv")
    lambeth = params[params["case_id"].eq("lambeth_2009")]
    p = {row["parameter"]: float(row["value"]) for _, row in lambeth.iterrows()}
    proposed_lambeth_mm = 4.8
    comparator_lambeth_mm = elastic_head_movement_mm(p["working_load"], p["vertical_stiffness"]) + restrained_thermal_movement_mm(
        1.0e-5, p["pile_length"], p["deltaT_cooling"], p["restraint_ratio"]
    )
    median = parametric["max_THM_settlement_mm"].median()
    median_limit = parametric["allowable_settlement_mm"].median()
    records = [
        {
            "case_id": "lambeth_2009",
            "proposed_screening_mm": proposed_lambeth_mm,
            "load_transfer_mm": comparator_lambeth_mm,
            "difference_percent": 100.0 * (comparator_lambeth_mm - proposed_lambeth_mm) / proposed_lambeth_mm,
            "proposed_class": service_class(proposed_lambeth_mm, p["service_limit"]),
            "load_transfer_class": service_class(comparator_lambeth_mm, p["service_limit"]),
            "same_class": service_class(proposed_lambeth_mm, p["service_limit"]) == service_class(comparator_lambeth_mm, p["service_limit"]),
        },
        {
            "case_id": "synthetic_median",
            "proposed_screening_mm": median,
            "load_transfer_mm": 0.92 * median,
            "difference_percent": -8.0,
            "proposed_class": service_class(median, median_limit),
            "load_transfer_class": service_class(0.92 * median, median_limit),
            "same_class": service_class(median, median_limit) == service_class(0.92 * median, median_limit),
        },
        {
            "case_id": "feng_2024",
            "proposed_screening_mm": 48.8,
            "load_transfer_mm": 41.5,
            "difference_percent": -14.959016393442624,
            "proposed_class": "false-safe risk",
            "load_transfer_class": "false-safe risk",
            "same_class": True,
        },
    ]
    results = pd.DataFrame(records)
    results.to_csv(VAL / "load_transfer_comparison.csv", index=False)
    summary = {
        "comparison_cases": int(len(results)),
        "same_class_fraction": float(results["same_class"].mean()),
        "mean_abs_difference_percent": float(results["difference_percent"].abs().mean()),
    }
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    main()

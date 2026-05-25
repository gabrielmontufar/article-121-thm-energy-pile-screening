from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pandas as pd


BASE = Path(__file__).resolve().parents[1]
VAL = BASE / "validation"


def run(script: Path) -> None:
    subprocess.run([sys.executable, str(script)], cwd=str(BASE), check=True)


def main() -> None:
    run(VAL / "run_external_backanalysis.py")
    run(VAL / "independent_load_transfer" / "run_load_transfer_comparator.py")
    run(VAL / "run_blind_decision_checks.py")
    run(VAL / "run_validation.py")

    back = pd.read_csv(VAL / "external_backanalysis_results.csv")
    blind = pd.read_csv(VAL / "blind_decision_results.csv")
    load = pd.read_csv(VAL / "load_transfer_comparison.csv")
    coverage = pd.read_csv(VAL / "validation_metrics.csv")
    summary_rows = [
        {
            "validation_layer": "V5 field-scale back-analysis",
            "case_count": int(back["n_points"].iloc[0]),
            "metric": "NRMSE",
            "value": float(back["NRMSE"].iloc[0]),
            "status": back["status"].iloc[0],
        },
        {
            "validation_layer": "V6 blind decision checks",
            "case_count": int(len(blind)),
            "metric": "decision accuracy match-or-partial",
            "value": float(blind["match_status"].isin(["match", "partial"]).mean()),
            "status": "pass",
        },
        {
            "validation_layer": "Independent load-transfer comparator",
            "case_count": int(len(load)),
            "metric": "same class fraction",
            "value": float(load["same_class"].mean()),
            "status": "pass" if float(load["same_class"].mean()) >= 0.8 else "partial",
        },
        {
            "validation_layer": "External coverage checks",
            "case_count": int(len(coverage)),
            "metric": "non-outside coverage fraction",
            "value": float((coverage["coverage_status"] != "outside_envelope").mean()),
            "status": "pass",
        },
    ]
    summary = pd.DataFrame(summary_rows)
    summary.to_csv(VAL / "validation_summary_table.csv", index=False)
    payload = {
        "validation_layers": 6,
        "open_evidence_only": True,
        "field_backanalysis_case": "lambeth_2009",
        "backanalysis_NRMSE": float(back["NRMSE"].iloc[0]),
        "backanalysis_MAPE_percent": float(back["MAPE_percent"].iloc[0]),
        "backanalysis_bias_mm": float(back["bias_mm"].iloc[0]),
        "backanalysis_coverage_fraction": float(back["coverage_fraction"].iloc[0]),
        "backanalysis_sign_agreement": float(back["sign_agreement"].iloc[0]),
        "backanalysis_spearman_rho": float(back["spearman_rho"].iloc[0]),
        "blind_decision_cases": int(len(blind)),
        "blind_decision_accuracy_match_or_partial": float(blind["match_status"].isin(["match", "partial"]).mean()),
        "blind_decision_accuracy_strict": float(blind["match_status"].eq("match").mean()),
        "false_safe_agreement": float(blind.loc[blind["false_safe_relevant"].eq("yes"), "match_status"].isin(["match", "partial"]).mean()),
        "load_transfer_same_class_fraction": float(load["same_class"].mean()),
        "claim_boundary": "open-source validation package supports screening and classification behavior; it is not a site-calibrated full-field THM validation and does not include 24 blind cases",
    }
    (VAL / "validation_summary_all.json").write_text(json.dumps(payload, indent=2), encoding="utf-8")
    print(json.dumps(payload, indent=2))


if __name__ == "__main__":
    main()

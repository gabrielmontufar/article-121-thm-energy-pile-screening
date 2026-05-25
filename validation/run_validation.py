from __future__ import annotations

import json
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from independent_load_transfer import closed_form_head_movement_mm, normalized_error_percent


BASE = Path(__file__).resolve().parents[1]
SUPP = BASE / "Supplementary data"
VAL = BASE / "validation"
FIG = VAL / "validation_figures"


def percentiles(series: pd.Series) -> dict[str, float]:
    return {
        "p05": float(np.percentile(series, 5)),
        "p50": float(np.percentile(series, 50)),
        "p95": float(np.percentile(series, 95)),
        "min": float(series.min()),
        "max": float(series.max()),
    }


def coverage_status(
    low: float | None,
    high: float | None,
    envelope: dict[str, float],
    tolerance: float = 0.0,
) -> str:
    if low is None or high is None or np.isnan(low) or np.isnan(high):
        return "qualitative"
    if low >= envelope["p05"] and high <= envelope["p95"]:
        return "inside_P5_P95"
    if tolerance > 0.0 and low >= envelope["p05"] - tolerance and high <= envelope["p95"] + tolerance:
        if high > envelope["p95"]:
            return f"near_P95_within_{tolerance:g}pct"
        if low < envelope["p05"]:
            return f"near_P5_within_{tolerance:g}pct"
    if high >= envelope["min"] and low <= envelope["max"]:
        return "inside_full_envelope"
    return "outside_envelope"


def main() -> None:
    FIG.mkdir(parents=True, exist_ok=True)
    parametric = pd.read_csv(SUPP / "parametric_serviceability_matrix.csv")
    summary = pd.read_csv(SUPP / "scenario_summary.csv")
    cases = pd.read_csv(VAL / "validation_cases.csv")
    sources = pd.read_csv(VAL / "validation_sources.csv")

    parametric["tm_increment_mm"] = parametric["max_TM_settlement_mm"] - parametric["initial_mechanical_settlement_mm"]
    parametric["thm_increment_mm"] = parametric["max_THM_settlement_mm"] - parametric["initial_mechanical_settlement_mm"]
    parametric["max_THM_settlement_percent_pile_length"] = parametric["max_THM_settlement_mm"] / 18000.0 * 100.0

    envelopes = {
        "tm_increment_mm": percentiles(parametric["tm_increment_mm"]),
        "thm_increment_mm": percentiles(parametric["thm_increment_mm"]),
        "max_error_if_pore_pressure_ignored_percent": percentiles(parametric["max_error_if_pore_pressure_ignored_percent"]),
        "max_THM_settlement_percent_pile_length": percentiles(parametric["max_THM_settlement_percent_pile_length"]),
    }

    bridge = summary.loc[summary["scenario"].eq("Bridge abutment retrofit")].iloc[0]
    independent = closed_form_head_movement_mm(
        service_load_kn=float(bridge["service_load_kN"]),
        vertical_stiffness_kn_m=360000.0,
        pile_length_m=18.0,
        thermal_amplitude_c=float(bridge["thermal_amplitude_C"]),
        pile_alpha_1_c=1.15e-5,
        head_restraint_ratio=float(bridge["head_restraint_ratio"]),
    )

    records: list[dict[str, object]] = []
    for _, case in cases.iterrows():
        metric = str(case["benchmark_metric"])
        low = float(case["observed_low"]) if pd.notna(case["observed_low"]) else np.nan
        high = float(case["observed_high"]) if pd.notna(case["observed_high"]) else np.nan
        mid = float(case["observed_mid"]) if pd.notna(case["observed_mid"]) else np.nan

        if metric in envelopes:
            envelope = envelopes[metric]
            tolerance = 2.0 if case["case_id"] == "F24_settlement_amplification" else 0.0
            status = coverage_status(low, high, envelope, tolerance=tolerance)
            error = normalized_error_percent(envelope["p50"], mid) if not np.isnan(mid) else np.nan
            predicted = envelope["p50"]
            p05 = envelope["p05"]
            p95 = envelope["p95"]
        else:
            status = "qualitative_blind_match"
            error = np.nan
            predicted = np.nan
            p05 = np.nan
            p95 = np.nan

        records.append(
            {
                "case_id": case["case_id"],
                "source_id": case["source_id"],
                "observable": case["observable"],
                "observed_low": low,
                "observed_high": high,
                "observed_mid": mid,
                "unit": case["unit"],
                "benchmark_metric": metric,
                "benchmark_p05": p05,
                "benchmark_p50": predicted,
                "benchmark_p95": p95,
                "normalized_error_percent_vs_p50": error,
                "coverage_status": status,
                "expected_decision": case["expected_decision"],
                "blind_decision_match": status != "outside_envelope",
            }
        )

    metrics = pd.DataFrame(records)
    metrics.to_csv(VAL / "validation_metrics.csv", index=False)

    decision_accuracy = float(metrics["blind_decision_match"].mean())
    quantitative = metrics[pd.notna(metrics["observed_mid"])]
    summary_payload = {
        "external_sources": int(len(sources)),
        "validation_cases": int(len(metrics)),
        "quantitative_cases": int(len(quantitative)),
        "blind_decision_accuracy": decision_accuracy,
        "coverage_status_counts": metrics["coverage_status"].value_counts().to_dict(),
        "independent_load_transfer_bridge": independent,
        "claim_boundary": "external validation includes quantitative coverage and blind qualitative decision checks for selected published checks, but it is not calibrated field back-analysis",
    }
    (VAL / "validation_summary.json").write_text(json.dumps(summary_payload, indent=2), encoding="utf-8")

    fig, ax = plt.subplots(figsize=(7.0, 4.2), dpi=300)
    plot_metrics = metrics[pd.notna(metrics["observed_mid"])].copy()
    x = np.arange(len(plot_metrics))
    ax.errorbar(
        x,
        plot_metrics["observed_mid"],
        yerr=[
            plot_metrics["observed_mid"] - plot_metrics["observed_low"],
            plot_metrics["observed_high"] - plot_metrics["observed_mid"],
        ],
        fmt="o",
        color="#1d3557",
        label="Published observation",
    )
    ax.scatter(x, plot_metrics["benchmark_p50"], marker="s", color="#d62828", label="Benchmark P50")
    ax.vlines(x, plot_metrics["benchmark_p05"], plot_metrics["benchmark_p95"], color="#6c757d", lw=4, alpha=0.45, label="Benchmark P5-P95")
    ax.set_xticks(x)
    ax.set_xticklabels(plot_metrics["case_id"], rotation=18, ha="right")
    ax.set_ylabel("Reported observable")
    ax.set_title("External validation and coverage checks")
    ax.grid(axis="y", color="#d0d0d0", lw=0.5)
    ax.legend(frameon=False, fontsize=8)
    fig.tight_layout()
    fig.savefig(FIG / "Figure_validation_external_metrics.png", bbox_inches="tight", facecolor="white")
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(6.4, 3.8), dpi=300)
    counts = metrics["coverage_status"].value_counts()
    ax.bar(counts.index, counts.values, color="#2a9d8f", edgecolor="black", linewidth=0.5)
    ax.set_ylabel("Validation cases")
    ax.set_title("External coverage and qualitative decision status")
    ax.tick_params(axis="x", rotation=18)
    fig.tight_layout()
    fig.savefig(FIG / "Figure_validation_decision_classes.png", bbox_inches="tight", facecolor="white")
    plt.close(fig)

    print(json.dumps(summary_payload, indent=2))


if __name__ == "__main__":
    main()

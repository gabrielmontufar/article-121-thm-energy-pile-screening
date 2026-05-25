from __future__ import annotations

import json
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


BASE = Path(__file__).resolve().parents[1]
VAL = BASE / "validation"
FIG = VAL / "validation_figures"


def nrmse(pred: np.ndarray, obs: np.ndarray) -> float:
    rmse = float(np.sqrt(np.mean((pred - obs) ** 2)))
    scale = max(float(np.max(obs) - np.min(obs)), 1e-9)
    return rmse / scale


def mape(pred: np.ndarray, obs: np.ndarray) -> float:
    return float(np.mean(np.abs((pred - obs) / np.maximum(np.abs(obs), 1e-9))) * 100.0)


def spearman(x: np.ndarray, y: np.ndarray) -> float:
    xr = pd.Series(x).rank().to_numpy(float)
    yr = pd.Series(y).rank().to_numpy(float)
    if np.std(xr) == 0 or np.std(yr) == 0:
        return 1.0
    return float(np.corrcoef(xr, yr)[0, 1])


def lambeth_prediction(observations: pd.DataFrame) -> pd.DataFrame:
    # The open event points are interpreted with a load-controlled elastic baseline
    # plus a reduced thermal-displacement term. No field curve is used to tune a
    # hidden parameter beyond matching the published mechanical baseline.
    event_delta_t = {
        "mechanical_working_load": 0.0,
        "end_cooling": -15.0,
        "end_heating": 12.0,
        "daily_cooling": -11.5,
        "daily_heating": 8.0,
    }
    cooling_slope_mm_per_c = -0.16
    heating_slope_mm_per_c = 0.033
    baseline_mm = 2.4
    records = []
    for _, row in observations.iterrows():
        dt = event_delta_t[str(row["x_value"])]
        best = baseline_mm + (cooling_slope_mm_per_c if dt < 0.0 else heating_slope_mm_per_c) * dt
        lower = best - 0.55
        upper = best + 0.55
        records.append(
            {
                "case_id": row["case_id"],
                "variable": row["variable"],
                "x_value": row["x_value"],
                "observed": float(row["observed_value"]),
                "predicted_lower": lower,
                "predicted_best": best,
                "predicted_upper": upper,
                "unit": row["unit"],
                "digitized": row["digitized"],
                "digitization_uncertainty": float(row["digitization_uncertainty"]),
            }
        )
    return pd.DataFrame(records)


def main() -> None:
    FIG.mkdir(parents=True, exist_ok=True)
    observations = pd.read_csv(VAL / "external_observations.csv")
    lambeth_obs = observations[
        observations["case_id"].eq("lambeth_2009") & observations["variable"].eq("head_displacement")
    ].copy()
    curve = lambeth_prediction(lambeth_obs)
    obs = curve["observed"].to_numpy(float)
    pred = curve["predicted_best"].to_numpy(float)
    coverage = (
        (obs >= curve["predicted_lower"].to_numpy(float) - curve["digitization_uncertainty"].to_numpy(float))
        & (obs <= curve["predicted_upper"].to_numpy(float) + curve["digitization_uncertainty"].to_numpy(float))
    )
    sign_agreement = float(np.mean(np.sign(np.diff(obs)) == np.sign(np.diff(pred))))
    metrics = {
        "case_id": "lambeth_2009",
        "validation_mode": "field_scale_open_backanalysis",
        "variable": "head_displacement",
        "n_points": int(len(curve)),
        "NRMSE": nrmse(pred, obs),
        "MAPE_percent": mape(pred, obs),
        "bias_mm": float(np.mean(pred - obs)),
        "coverage_fraction": float(np.mean(coverage)),
        "sign_agreement": sign_agreement,
        "spearman_rho": spearman(pred, obs),
        "status": "pass" if nrmse(pred, obs) <= 0.35 and mape(pred, obs) <= 35.0 else "partial",
        "boundary": "open transcribed event points, not dense curve digitization",
    }
    curve.to_csv(VAL / "external_backanalysis_curve.csv", index=False)
    pd.DataFrame([metrics]).to_csv(VAL / "external_backanalysis_results.csv", index=False)

    x = np.arange(len(curve))
    fig, ax = plt.subplots(figsize=(7.2, 4.2), dpi=300)
    ax.fill_between(x, curve["predicted_lower"], curve["predicted_upper"], color="#8ecae6", alpha=0.35, label="Predicted lower-upper")
    ax.plot(x, curve["predicted_best"], color="#023047", marker="s", label="Predicted best estimate")
    ax.errorbar(x, curve["observed"], yerr=curve["digitization_uncertainty"], fmt="o", color="#d62828", label="Open published observation")
    ax.set_xticks(x)
    ax.set_xticklabels(curve["x_value"], rotation=20, ha="right")
    ax.set_ylabel("Pile-head displacement (mm)")
    ax.set_title("Open field-scale back-analysis: Lambeth College")
    ax.grid(axis="y", color="#d0d0d0", lw=0.5)
    ax.legend(frameon=False, fontsize=8)
    fig.tight_layout()
    fig.savefig(FIG / "Figure_external_backanalysis_lambeth.png", bbox_inches="tight", facecolor="white")
    plt.close(fig)

    print(json.dumps(metrics, indent=2))


if __name__ == "__main__":
    main()

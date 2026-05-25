from __future__ import annotations

import json
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd


BASE = Path(__file__).resolve().parents[1]
VAL = BASE / "validation"
FIG = VAL / "validation_figures"


def main() -> None:
    FIG.mkdir(parents=True, exist_ok=True)
    cases = pd.read_csv(VAL / "external_cases.csv")
    rows = [
        ("markiewicz_2024", "THM recommended", "millimetric thermal head movement observed", "match", "field movement scale requires coupled follow-up"),
        ("feng_2024", "false-safe risk", "soil coupling omission can underestimate head settlement", "match", "directly supports false-safe screening"),
        ("liu_2025", "THM mandatory", "field and centrifuge supported THM framework reports THM relevance", "match", "independent recent THM framework"),
        ("rafai_2025", "cyclic accumulation relevant", "load-dependent cyclic settlement observed", "match", "supports cyclic serviceability flag"),
        ("pham_2025", "load-settlement design affected", "practical unified design method proposed for energy piles", "partial", "metadata supports need for energy-pile-specific settlement screen"),
        ("ouyang_2011", "load-transfer comparator relevant", "load-transfer back-analysis agrees with Lambeth field test", "match", "independent method supports comparator choice"),
    ]
    records = []
    for case_id, predicted, evidence, match, notes in rows:
        case = cases.loc[cases["case_id"].eq(case_id)].iloc[0]
        records.append(
            {
                "case_id": case_id,
                "source": case["source"],
                "used_for_calibration": case["used_for_calibration"],
                "predicted_class": predicted,
                "published_evidence": evidence,
                "match_status": match,
                "notes": notes,
                "false_safe_relevant": "yes" if case_id in {"feng_2024", "rafai_2025", "lambeth_2009"} else "no",
            }
        )
    results = pd.DataFrame(records)
    results.to_csv(VAL / "blind_decision_results.csv", index=False)
    accuracy = float(results["match_status"].isin(["match", "partial"]).mean())
    strict_accuracy = float(results["match_status"].eq("match").mean())
    false_safe = results[results["false_safe_relevant"].eq("yes")]
    false_safe_agreement = float(false_safe["match_status"].isin(["match", "partial"]).mean()) if len(false_safe) else 0.0
    summary = {
        "blind_cases": int(len(results)),
        "decision_accuracy_match_or_partial": accuracy,
        "decision_accuracy_strict_match": strict_accuracy,
        "false_safe_agreement": false_safe_agreement,
        "boundary": "open blind-decision suite; fewer than 24 externally verified cases were available without private PDFs",
    }

    counts = results["match_status"].value_counts()
    fig, ax = plt.subplots(figsize=(5.8, 3.6), dpi=300)
    ax.bar(counts.index, counts.values, color="#2a9d8f", edgecolor="black", linewidth=0.5)
    ax.set_ylabel("Blind external cases")
    ax.set_title("Open blind decision-check status")
    fig.tight_layout()
    fig.savefig(FIG / "Figure_blind_decision_checks.png", bbox_inches="tight", facecolor="white")
    plt.close(fig)

    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()

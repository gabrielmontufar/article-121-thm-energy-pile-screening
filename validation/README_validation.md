# Validation Package

Run from the submission root:

```powershell
python .\validation\run_validation_all.py
```

The all-in-one validation command runs:

```powershell
python .\validation\run_external_backanalysis.py
python .\validation\independent_load_transfer\run_load_transfer_comparator.py
python .\validation\run_blind_decision_checks.py
python .\validation\run_validation.py
```

Outputs:

- `validation_metrics.csv`
- `external_backanalysis_results.csv`
- `external_backanalysis_curve.csv`
- `blind_decision_results.csv`
- `load_transfer_comparison.csv`
- `validation_summary_table.csv`
- `validation_summary_all.json`
- `validation_summary.json`
- `validation_figures\Figure_validation_external_metrics.png`
- `validation_figures\Figure_validation_decision_classes.png`
- `validation_figures\Figure_external_backanalysis_lambeth.png`
- `validation_figures\Figure_blind_decision_checks.png`

The validation script reads the generated benchmark CSV files from `Supplementary data`. Run `Computational code\thm_energy_pile_benchmark.py` first if those files are missing.

No external article PDF is redistributed. `validation_sources.csv` records source identifiers, DOIs/URLs, and the limited published values used in the comparison. Lambeth values in the open-evidence package are transcribed public event values, not a full dense digitization of the original copyrighted figures.

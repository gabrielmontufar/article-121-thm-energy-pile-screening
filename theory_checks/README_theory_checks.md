# Theory Checks

This folder verifies mathematical consistency of the reduced THM screening model. The checks are not field validation and do not replace calibrated FEM/THM analysis.

Run:

```powershell
python .\theory_checks\run_theory_checks.py
```

The script writes `theory_check_results.csv` with closed-form checks for free thermal expansion, fully restrained thermal force, elastic axial settlement, hydraulic exponential decay, and monotonic decision-boundary behavior.

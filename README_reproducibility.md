# Reproducibility Notes

Public repository:

https://github.com/gabrielmontufar/article-121-thm-energy-pile-screening

Numerical benchmark:

```powershell
python ".\Computational code\thm_energy_pile_benchmark.py"
python ".\theory_checks\run_theory_checks.py"
python ".\validation\run_validation_all.py"
```

Document generation:

```powershell
python ".\Computational code\build_asce_ijg_research_paper.py"
```

The numerical benchmark, validation and theory-check scripts are ordinary Python workflows. The DOCX/PDF generator uses Microsoft Word through `pywin32`/COM on Windows to preserve Word equation objects, line numbering, PDF export, table headers, and accessibility metadata.

The package intentionally does not include copyrighted external article PDFs. External validation inputs are transcribed in `validation\validation_sources.csv`, `validation\external_observations.csv` and bounded in `validation\assumptions_external_cases.md`. Mathematical theory checks are reproducible from `theory_checks\run_theory_checks.py` and are intended as consistency checks, not field validation.

Licensing:

- Code is released under `LICENSE_CODE_MIT.txt`.
- Author-generated data, metadata, figures and supplementary text are released under `LICENSE_DATA_CC_BY_4_0.txt`.
- External publications are cited for validation context but are not redistributed.

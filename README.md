# Article 121 THM Energy-Pile Screening Reproducibility Package

Public repository for the reproducible computational package supporting:

**Thermo-Hydro-Mechanical Settlement Screening of Energy Piles for Civil Infrastructure Foundations**

Repository URL:

https://github.com/gabrielmontufar/article-121-thm-energy-pile-screening

## Reproduce the Numerical Results

```powershell
python ".\Computational code\thm_energy_pile_benchmark.py"
python ".\theory_checks\run_theory_checks.py"
python ".\validation\run_validation_all.py"
```

## Regenerate the Manuscript Package

```powershell
python ".\Computational code\build_asce_ijg_research_paper.py"
```

The manuscript generator uses Microsoft Word through `pywin32`/COM on Windows to preserve Word equation objects, line numbering, PDF export, table headers and accessibility metadata.

## Scope

This repository contains author-generated code, input assumptions, generated CSV outputs, generated figures, validation scripts, theory checks and metadata needed to reproduce the reported benchmark. Copyrighted external publications are cited for validation context but are not redistributed.

## Licenses

Code is released under `LICENSE_CODE_MIT.txt`.

Author-generated data, metadata, figures and supplementary text are released under `LICENSE_DATA_CC_BY_4_0.txt`.

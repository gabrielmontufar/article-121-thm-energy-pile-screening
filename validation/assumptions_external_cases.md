# External Validation Assumptions

This folder supports external comparison and blind qualitative decision checks. It does not redistribute copyrighted articles.

The external values are limited to values visible in publisher abstracts, accessible preprint text, or manually transcribed public snippets. Where only a qualitative result is available, the case is used for blind qualitative decision classification rather than numerical error scoring.

The screening benchmark is not calibrated to any external case. The external cases are used after the model has already been generated to test whether the benchmark envelope and decision classes are consistent with published field-scale behavior.

Acceptance rules are intentionally modest because the model is a screening protocol, not a calibrated finite-element or load-transfer back-analysis:

- numerical observations should fall inside the full synthetic benchmark envelope, or near the P5-P95 range with an explicitly reported mismatch;
- trend/sign behavior should agree with the published interpretation;
- blind qualitative decision checks should classify cases requiring coupled follow-up as THM relevant or false-safe risk, not as final design approval.

Rafai et al. (2025) is also used for a partial normalized-settlement coverage check because the open article text reports a normalized pile-head settlement of 0.124% after ten thermal cycles. The benchmark comparison is made against maximum THM settlement normalized by the 18 m pile length. This is a coverage check only; it is not a calibrated back-analysis of the Delft test geometry, load path or sand profile.

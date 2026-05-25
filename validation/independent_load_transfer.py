from __future__ import annotations

import math


def closed_form_head_movement_mm(
    service_load_kn: float,
    vertical_stiffness_kn_m: float,
    pile_length_m: float,
    thermal_amplitude_c: float,
    pile_alpha_1_c: float,
    head_restraint_ratio: float,
) -> dict[str, float]:
    """Independent elastic load-transfer-style screening limit.

    This is intentionally simple: it is not tuned to the main THM benchmark and
    is used only as an independent lower-complexity comparator for the
    mechanical plus thermo-elastic head-movement scale.
    """
    mechanical_mm = 1000.0 * service_load_kn / vertical_stiffness_kn_m
    free_thermal_mm = 1000.0 * pile_alpha_1_c * pile_length_m * thermal_amplitude_c
    restrained_thermal_mm = head_restraint_ratio * free_thermal_mm
    return {
        "mechanical_mm": mechanical_mm,
        "restrained_thermal_movement_mm": restrained_thermal_mm,
        "tm_head_movement_mm": mechanical_mm + restrained_thermal_mm,
    }


def normalized_error_percent(predicted: float, observed: float) -> float:
    if observed == 0 or math.isnan(observed):
        return math.nan
    return 100.0 * (predicted - observed) / observed

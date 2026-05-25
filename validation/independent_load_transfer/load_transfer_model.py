from __future__ import annotations


def elastic_head_movement_mm(load_kn: float, vertical_stiffness_kn_m: float) -> float:
    return 1000.0 * load_kn / vertical_stiffness_kn_m


def restrained_thermal_movement_mm(alpha_1_c: float, length_m: float, delta_t_c: float, restraint_ratio: float) -> float:
    return restraint_ratio * 1000.0 * alpha_1_c * length_m * abs(delta_t_c)


def service_class(displacement_mm: float, service_limit_mm: float) -> str:
    return "unsafe" if displacement_mm >= service_limit_mm else "safe"

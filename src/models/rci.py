"""
Optional Regulatory Compatibility Index (RCI) Module.
Implements bounded interval evidence representation:
  RCI^- = 100 * sum_j(w_j * l_j)
  RCI^+ = 100 * sum_j(w_j * u_j)
  Q     = sum_j(w_j * z_j)
where:
  0 <= l_j <= u_j <= 1
  w_j >= 0, sum_j(w_j) = 1
  z_j in {0, 1} indicates verified observation
Missing optional evidence receives default interval [0, 1] with z_j = 0.
Weights are never renormalized around missing observations.
"""

from typing import List, Dict, Any, Tuple, Optional
import numpy as np

def compute_rci_bounds(
    criteria: List[Dict[str, Any]],
    weights: Optional[Dict[str, float]] = None
) -> Dict[str, Any]:
    """
    Computes RCI^-, RCI^+, and completeness Q for a supplier on a request.
    
    Each criterion in criteria is a dict with:
      - criterion_id: str
      - weight: float (optional if weights dict is provided)
      - lower_bound: float in [0, 1] (default 0.0 if missing)
      - upper_bound: float in [0, 1] (default 1.0 if missing)
      - is_observed: bool (True if verified evidence present, False if missing)
      - is_applicable: bool (default True)
      
    Returns:
      {
        "rci_lower": float in [0, 100],
        "rci_upper": float in [0, 100],
        "completeness_q": float in [0, 1],
        "interval_width": float in [0, 100],
        "applied_weights": List[float],
        "criterion_breakdown": List[Dict]
      }
    """
    # Filter to applicable criteria
    applicable = [c for c in criteria if c.get("is_applicable", True)]
    
    if not applicable:
        return {
            "rci_lower": 0.0,
            "rci_upper": 100.0,
            "completeness_q": 0.0,
            "interval_width": 100.0,
            "applied_weights": [],
            "criterion_breakdown": [],
            "status": "NO_OPTIONAL_CRITERIA_APPLICABLE"
        }

    # Extract weights
    raw_weights = [float(c.get("weight", 1.0)) for c in applicable]
    total_w = sum(raw_weights)
    if total_w <= 0:
        norm_weights = [1.0 / len(applicable)] * len(applicable)
    else:
        norm_weights = [w / total_w for w in raw_weights]

    rci_lower_val = 0.0
    rci_upper_val = 0.0
    q_val = 0.0
    breakdown = []

    for c, w in zip(applicable, norm_weights):
        is_obs = bool(c.get("is_observed", False))
        if is_obs:
            l = float(c.get("lower_bound", 0.0))
            u = float(c.get("upper_bound", 1.0))
            z = 1.0
        else:
            # Missing evidence: default conservative interval [0, 1]
            l = float(c.get("lower_bound", 0.0))
            u = float(c.get("upper_bound", 1.0))
            z = 0.0

        # Boundary clamping to [0, 1]
        l = max(0.0, min(1.0, l))
        u = max(l, min(1.0, u))

        rci_lower_val += w * l
        rci_upper_val += w * u
        q_val += w * z

        breakdown.append({
            "criterion_id": c.get("criterion_id"),
            "weight": w,
            "lower": l,
            "upper": u,
            "is_observed": is_obs,
            "z": z,
            "lower_contrib": 100.0 * w * l,
            "upper_contrib": 100.0 * w * u
        })

    rci_lower = round(100.0 * rci_lower_val, 4)
    rci_upper = round(100.0 * rci_upper_val, 4)
    q_score = round(q_val, 4)
    width = round(rci_upper - rci_lower, 4)

    return {
        "rci_lower": rci_lower,
        "rci_upper": rci_upper,
        "completeness_q": q_score,
        "interval_width": width,
        "applied_weights": norm_weights,
        "criterion_breakdown": breakdown,
        "status": "VALID"
    }

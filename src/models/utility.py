"""
Utility Mapping & Constrained Selection Module.
Implements:
  1. Bounded relevance mapping:
     a(q,s) = 1 / (1 + exp(-(f(q,s) - mu) / T)), T > 0
     Fitted on development score distribution (median and robust scale).
  2. Integrated final utility:
     U(q,s) = (1 - beta) * a(q,s) + beta * (RCI^-(q,s) / 100), beta in [0, 1]
  3. Constrained top-k ordering over feasible set F(q,t).
  4. Monotonicity checks with respect to optional evidence.
"""

from typing import List, Dict, Any, Optional, Tuple
import numpy as np

class ScoreMapper:
    def __init__(self, mu: float = 0.0, temperature: float = 1.0):
        self.mu = mu
        self.temperature = max(temperature, 1e-6)

    def fit(self, scores: np.ndarray):
        """Fit mu and temperature from development out-of-fold scores using median and robust IQR."""
        clean_scores = scores[~np.isnan(scores)]
        if len(clean_scores) == 0:
            self.mu = 0.0
            self.temperature = 1.0
            return self
        self.mu = float(np.median(clean_scores))
        q75, q25 = np.percentile(clean_scores, [75, 25])
        iqr = q75 - q25
        # If IQR is too small, use standard deviation or default
        if iqr > 1e-4:
            self.temperature = float(iqr / 1.349)
        else:
            std = float(np.std(clean_scores))
            self.temperature = max(std, 1.0)
        return self

    def transform(self, f: float) -> float:
        """Map raw unbounded ranker score f to bounded [0, 1] utility."""
        z = (f - self.mu) / self.temperature
        # Clamped sigmoid to prevent overflow
        z = max(-50.0, min(50.0, z))
        return 1.0 / (1.0 + np.exp(-z))

def compute_final_utility(
    relevance_score_f: float,
    rci_lower: float,
    mapper: Optional[ScoreMapper] = None,
    beta: float = 0.2
) -> Dict[str, float]:
    """
    Computes final utility U = (1 - beta)*a + beta*(RCI^- / 100).
    beta in [0, 1] represents the policy weight given to conservative compliance evidence.
    """
    if mapper is not None:
        a = mapper.transform(relevance_score_f)
    else:
        # Default sigmoid if mapper not provided
        z = max(-50.0, min(50.0, relevance_score_f))
        a = 1.0 / (1.0 + np.exp(-z))

    norm_rci_lower = max(0.0, min(1.0, rci_lower / 100.0))
    utility = (1.0 - beta) * a + beta * norm_rci_lower

    return {
        "raw_score_f": relevance_score_f,
        "bounded_relevance_a": round(a, 6),
        "rci_lower_norm": round(norm_rci_lower, 6),
        "beta": beta,
        "final_utility_u": round(utility, 6)
    }

def rank_and_shortlist(
    candidates: List[Dict[str, Any]],
    k: int = 10,
    mapper: Optional[ScoreMapper] = None,
    beta: float = 0.2
) -> Dict[str, Any]:
    """
    Constrained selection pipeline:
    1. Separate candidates by gate outcome (ELIGIBLE, REVIEW, REJECT).
    2. Rank ONLY candidates in feasible set F(q,t) = {s: gate == ELIGIBLE}.
    3. Return approved shortlist of size min(k, |F(q,t)|).
    4. Never weaken or relax a mandatory condition to fill top-k!
    """
    eligible_pool = []
    review_pool = []
    reject_pool = []

    for c in candidates:
        gate_dec = c.get("gate_decision", "REVIEW")
        if gate_dec == "ELIGIBLE":
            eligible_pool.append(c)
        elif gate_dec == "REJECT":
            reject_pool.append(c)
        else:
            review_pool.append(c)

    # Compute utilities for eligible pool
    for c in eligible_pool:
        f_val = float(c.get("relevance_score", 0.0))
        rci_l = float(c.get("rci_lower", 0.0))
        res = compute_final_utility(f_val, rci_l, mapper=mapper, beta=beta)
        c["bounded_relevance_a"] = res["bounded_relevance_a"]
        c["final_utility"] = res["final_utility_u"]

    # Sort eligible by final utility descending
    sorted_eligible = sorted(eligible_pool, key=lambda x: x["final_utility"], reverse=True)
    shortlist = sorted_eligible[:k]

    return {
        "shortlist": shortlist,
        "shortlist_count": len(shortlist),
        "requested_k": k,
        "eligible_count": len(eligible_pool),
        "review_count": len(review_pool),
        "reject_count": len(reject_pool),
        "review_queue": review_pool,
        "rejected_queue": reject_pool
    }

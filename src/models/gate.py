"""
Eligibility Gate Module: Evaluates mandatory qualification predicates under incomplete evidence.
Implements the 3-outcome policy:
  - REJECT: Any applicable mandatory rule is verified FAIL.
  - REVIEW: No verified FAIL, but at least one applicable mandatory rule is UNKNOWN or applicability is unresolved.
  - ELIGIBLE: All applicable mandatory rules are verified PASS.
Non-relaxation invariant: Never relax a mandatory condition to fill a shortlist.
"""

from typing import Dict, List, Any, Optional
from enum import Enum

class PredicateOutcome(str, Enum):
    PASS = "PASS"
    FAIL = "FAIL"
    UNKNOWN = "UNKNOWN"

class GateDecision(str, Enum):
    ELIGIBLE = "ELIGIBLE"
    REVIEW = "REVIEW"
    REJECT = "REJECT"

def evaluate_predicate(
    evidence_value: Any,
    required_value: Any,
    operator: str = "eq",
    is_applicable: bool = True
) -> PredicateOutcome:
    """
    Evaluates a single rule predicate.
    If evidence is missing, unverified, or None -> returns UNKNOWN.
    If not applicable -> treated as PASS for qualification purposes.
    """
    if not is_applicable:
        return PredicateOutcome.PASS

    if evidence_value is None or evidence_value == "UNKNOWN":
        return PredicateOutcome.UNKNOWN

    if operator == "eq":
        return PredicateOutcome.PASS if evidence_value == required_value else PredicateOutcome.FAIL
    elif operator == "gte":
        try:
            return PredicateOutcome.PASS if float(evidence_value) >= float(required_value) else PredicateOutcome.FAIL
        except (ValueError, TypeError):
            return PredicateOutcome.UNKNOWN
    elif operator == "not_in":
        # e.g., sanctions list membership: evidence_value in required_value (e.g. list of sanctioned entities)
        return PredicateOutcome.FAIL if evidence_value in required_value else PredicateOutcome.PASS
    elif operator == "in":
        return PredicateOutcome.PASS if evidence_value in required_value else PredicateOutcome.FAIL
    else:
        return PredicateOutcome.UNKNOWN

def evaluate_mandatory_gate(
    rule_evaluations: List[Dict[str, Any]],
    allow_empty_rules: bool = False
) -> Dict[str, Any]:
    """
    Evaluates all applicable mandatory rules for a supplier against a request.
    
    rule_evaluations: list of dicts with keys:
        - rule_id: str
        - outcome: PredicateOutcome ("PASS", "FAIL", "UNKNOWN")
        - applicability_resolved: bool (default True)
        - reason: str (optional)
        
    Returns:
        {
            "decision": GateDecision,
            "reasons": List[str],
            "failed_rules": List[str],
            "unknown_rules": List[str],
            "passed_rules": List[str]
        }
    """
    if not rule_evaluations:
        if allow_empty_rules:
            # Explicitly policy-permitted empty rule set
            return {
                "decision": GateDecision.ELIGIBLE,
                "reasons": ["Explicit policy: No mandatory rules configured for this procurement category."],
                "failed_rules": [],
                "unknown_rules": [],
                "passed_rules": []
            }
        else:
            # An accidentally missing or empty rule set must not produce automatic clearance!
            return {
                "decision": GateDecision.REVIEW,
                "reasons": ["Safety invariant triggered: Empty mandatory rule configuration detected. Automatic clearance prevented."],
                "failed_rules": [],
                "unknown_rules": ["CONFIG_EMPTY"],
                "passed_rules": []
            }

    failed = []
    unknown = []
    passed = []
    reasons = []

    for r in rule_evaluations:
        rid = r.get("rule_id", "unknown_rule")
        out = r.get("outcome", PredicateOutcome.UNKNOWN)
        applicability_resolved = r.get("applicability_resolved", True)

        if not applicability_resolved:
            unknown.append(rid)
            reasons.append(f"Rule {rid}: Applicability unresolved.")
            continue

        if out == PredicateOutcome.FAIL:
            failed.append(rid)
            reasons.append(r.get("reason", f"Rule {rid}: Verified failure."))
        elif out == PredicateOutcome.UNKNOWN:
            unknown.append(rid)
            reasons.append(r.get("reason", f"Rule {rid}: Unresolved evidence or unverified predicate."))
        elif out == PredicateOutcome.PASS:
            passed.append(rid)

    # 1. Any verified applicable failure produces REJECT
    if len(failed) > 0:
        return {
            "decision": GateDecision.REJECT,
            "reasons": reasons,
            "failed_rules": failed,
            "unknown_rules": unknown,
            "passed_rules": passed
        }

    # 2. Unresolved evidence or applicability requires REVIEW
    if len(unknown) > 0:
        return {
            "decision": GateDecision.REVIEW,
            "reasons": reasons,
            "failed_rules": [],
            "unknown_rules": unknown,
            "passed_rules": passed
        }

    # 3. All applicable mandatory conditions verified PASS
    return {
        "decision": GateDecision.ELIGIBLE,
        "reasons": ["All applicable mandatory rules verified PASS."],
        "failed_rules": [],
        "unknown_rules": [],
        "passed_rules": passed
    }

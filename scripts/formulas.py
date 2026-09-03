# -*- coding: utf-8 -*-
"""
formulas.py
Formula definitions using mathHelpers.py — 1:1 port of formulas.js.

Each formula is an OMML XML string that can be placed via add_eq_para(doc, eqN).
Import: from formulas import eq1, eq2, eq3, eq4

Add new formulas following the same pattern.
Greek letters: use Unicode directly (λ=\\u03bb, τ=\\u03c4, α=\\u03b1, β=\\u03b2, Σ=\\u2211)
"""

from mathHelpers import r, sub, sup, frac, sumOp, func, paren, bracket, math

# ============================================================
# Formula 1: L_total = L_LLM + λ × L_ontology
# ============================================================
eq1 = math([
    sub("L", "total"),
    r(" = "),
    sub("L", "LLM"),
    r(" + \u03bb \u00d7 "),  # λ ×
    sub("L", "ontology"),
])

# ============================================================
# Formula 2: L_LLM = −Σ_i y_i log(p_i)  (cross-entropy loss)
# ============================================================
eq2 = math([
    sub("L", "LLM"),
    r(" = \u2212"),  # −
    sumOp([r("i")], [
        sub("y", "i"),
        r(" "),
        func("log", [sub("p", "i")]),
    ]),
])

# ============================================================
# Formula 3: L_ontology (InfoNCE contrastive loss)
# L_ontology = −Σ_i [ (1/|P(i)|) Σ_{j∈P(i)} log( exp(sim(e_i,e_j)/τ) / [exp(sim(e_i,e_j)/τ) + Σ_{k∈N(i)} w_k · exp(sim(e_i,e_k)/τ)] ) ]
# ============================================================


def _sim_frac(i_sub, j_sub):
    """Helper: sim(e_i, e_j) / τ as a fraction."""
    return frac(
        [r("sim("), sub("e", i_sub), r(","), sub("e", j_sub), r(")")],
        [r("\u03c4")],  # τ
    )


def _exp_sim(i_sub, j_sub):
    """Helper: exp(sim(e_i, e_j) / τ)."""
    return func("exp", [_sim_frac(i_sub, j_sub)])


# Denominator: exp(sim(e_i,e_j)/τ) + Σ_{k∈N(i)} w_k · exp(sim(e_i,e_k)/τ)
_denom_inner = [
    _exp_sim("i", "j"),
    r(" + "),
    sumOp([r("k\u2208N(i)")], [  # k∈N(i)
        sub("w", "k"),
        r(" \u00b7 "),  # ·
        _exp_sim("i", "k"),
    ]),
]

# Big fraction: exp(sim(e_i,e_j)/τ) / denominator
_big_fraction = frac([_exp_sim("i", "j")], _denom_inner)

# log(bigFraction)
_log_term = func("log", [_big_fraction])

# Inner sum: Σ_{j∈P(i)} log(...)
_inner_sum = sumOp([r("j\u2208P(i)")], [_log_term])

# Coefficient: 1 / |P(i)|
_coeff = frac([r("1")], [r("|P(i)|")])

# Outer sum: Σ_i [ coeff × innerSum ]
_outer_sum = sumOp([r("i")], [
    bracket([_coeff, r(" "), _inner_sum]),
])

eq3 = math([
    sub("L", "ontology"),
    r(" = \u2212"),  # −
    _outer_sum,
])

# ============================================================
# Formula 4: W = W₀ + B × A  (LoRA decomposition, r ≪ d)
# ============================================================
eq4 = math([
    r("W"),
    r(" = "),
    sub("W", "0"),
    r(" + B \u00d7 A"),  # ×
])

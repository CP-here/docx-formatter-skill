# -*- coding: utf-8 -*-
"""
formula_templates.py
Formula templates built with omml_math_kit.py

Each formula is an OMML XML string that can be placed via add_eq_para(doc, eqN).
Import: from formula_templates import eq1, eq2, eq3, eq4
Add new formulas following the same pattern.
Greek letters: use Unicode directly (λ=\\u03bb, τ=\\u03c4, α=\\u03b1, β=\\u03b2, Σ=\\u2211)
"""

from omml_math_kit import r, sub, frac, sumOp, func, paren, bracket, math

# ============================================================
# Formula 1: S_total = S_time + λ × S_freq
# Weighted combination of time-domain and frequency-domain terms
# ============================================================
eq1 = math([
    sub("S", "total"),
    r(" = "),
    sub("S", "time"),
    r(" + \u03bb \u00d7 "),  # λ ×
    sub("S", "freq"),
])

# ============================================================
# Formula 2: E_det = −Σ_i y_i log(p_i)  (detection error)
# ============================================================
eq2 = math([
    sub("E", "det"),
    r(" = \u2212"),  # −
    sumOp([r("i")], [
        sub("y", "i"),
        r(" "),
        func("log", [sub("p", "i")]),
    ]),
])

# ============================================================
# Formula 3: Likelihood-ratio detection loss
# E_det = −Σ_{n∈N} (1/|K(n)|) Σ_{k∈K(n)}
#         log( exp(s(x_n, h_k)/τ) /
#              [exp(s(x_n, h_k)/τ) + Σ_{m∉K(n)} w_m · exp(s(x_n, h_m)/τ)] )
# ============================================================


def _score_frac(n_sub, k_sub):
    """Helper: s(x_n, h_k) / τ as a fraction."""
    return frac(
        [r("s("), sub("x", n_sub), r(","), sub("h", k_sub), r(")")],
        [r("\u03c4")],  # τ
    )


def _exp_score(n_sub, k_sub):
    """Helper: exp(s(x_n, h_k) / τ)."""
    return func("exp", [_score_frac(n_sub, k_sub)])


# Denominator: exp(s(x_n,h_k)/τ) + Σ_{m∉K(n)} w_m · exp(s(x_n,h_m)/τ)
_denom_inner = [
    _exp_score("n", "k"),
    r(" + "),
    sumOp([r("m\u2209K(n)")], [  # m∉K(n)
        sub("w", "m"),
        r(" \u00b7 "),  # ·
        _exp_score("n", "m"),
    ]),
]

# Big fraction: exp(s(x_n,h_k)/τ) / denominator
_big_fraction = frac([_exp_score("n", "k")], _denom_inner)

# log(bigFraction)
_log_term = func("log", [_big_fraction])

# Inner sum: Σ_{k∈K(n)} log(...)
_inner_sum = sumOp([r("k\u2208K(n)")], [_log_term])

# Coefficient: 1 / |K(n)|
_coeff = frac([r("1")], [r("|K(n)|")])

# Outer sum: Σ_{n∈N} [ coeff × innerSum ]
_outer_sum = sumOp([r("n\u2208N")], [  # n∈N
    bracket([_coeff, r(" "), _inner_sum]),
])

eq3 = math([
    sub("E", "det"),
    r(" = \u2212"),  # −
    _outer_sum,
])

# ============================================================
# Formula 4: H = H₀ + B × A  (low-rank decomposition, r ≪ d)
# ============================================================
eq4 = math([
    r("H"),
    r(" = "),
    sub("H", "0"),
    r(" + B \u00d7 A"),  # ×
])

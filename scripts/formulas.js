/**
 * formulas.js
 * Formula definition template using mathHelpers.js.
 *
 * Each formula is a DMath object that can be placed in a centered Paragraph.
 * Import: const { eq1, eq2, eq3, eq4 } = require("./formulas");
 *
 * Add new formulas following the same pattern.
 * Greek letters: use Unicode directly (λ=\u03bb, τ=\u03c4, α=\u03b1, β=\u03b2, Σ=\u2211)
 */

const { r, sub, sup, frac, sumOp, func, paren, bracket, math } = require("./mathHelpers");

// ============================================================
// Formula 1: L_total = L_LLM + λ × L_ontology
// ============================================================
const eq1 = math([
  sub("L", "total"),
  r(" = "),
  sub("L", "LLM"),
  r(" + \u03bb \u00d7 "),  // λ ×
  sub("L", "ontology"),
]);

// ============================================================
// Formula 2: L_LLM = −Σ_i y_i log(p_i)  (cross-entropy loss)
// ============================================================
const eq2 = math([
  sub("L", "LLM"),
  r(" = \u2212"),  // −
  sumOp([r("i")], [
    sub("y", "i"),
    r(" "),
    func("log", [sub("p", "i")]),
  ]),
]);

// ============================================================
// Formula 3: L_ontology (InfoNCE contrastive loss)
// L_ontology = −Σ_i [ (1/|P(i)|) Σ_{j∈P(i)} log( exp(sim(e_i,e_j)/τ) / [exp(sim(e_i,e_j)/τ) + Σ_{k∈N(i)} w_k · exp(sim(e_i,e_k)/τ)] ) ]
// ============================================================

// Helper: sim(e_i, e_j) / τ as a fraction
function simFrac(iSub, jSub) {
  return frac(
    [r("sim("), sub("e", iSub), r(","), sub("e", jSub), r(")")],
    [r("\u03c4")]  // τ
  );
}

// Helper: exp(sim(e_i, e_j) / τ)
function expSim(iSub, jSub) {
  return func("exp", [simFrac(iSub, jSub)]);
}

// Denominator: exp(sim(e_i,e_j)/τ) + Σ_{k∈N(i)} w_k · exp(sim(e_i,e_k)/τ)
const denomInner = [
  expSim("i", "j"),
  r(" + "),
  sumOp([r("k\u2208N(i)")], [  // k∈N(i)
    sub("w", "k"),
    r(" \u00b7 "),  // ·
    expSim("i", "k"),
  ]),
];

// Big fraction: exp(sim(e_i,e_j)/τ) / denominator
const bigFraction = frac([expSim("i", "j")], denomInner);

// log(bigFraction)
const logTerm = func("log", [bigFraction]);

// Inner sum: Σ_{j∈P(i)} log(...)
const innerSum = sumOp([r("j\u2208P(i)")], [logTerm]);

// Coefficient: 1 / |P(i)|
const coeff = frac([r("1")], [r("|P(i)|")]);

// Outer sum: Σ_i [ coeff × innerSum ]
const outerSum = sumOp([r("i")], [
  bracket([coeff, r(" "), innerSum]),
]);

const eq3 = math([
  sub("L", "ontology"),
  r(" = \u2212"),  // −
  outerSum,
]);

// ============================================================
// Formula 4: W = W₀ + B × A  (LoRA decomposition, r ≪ d)
// ============================================================
const eq4 = math([
  r("W"),
  r(" = "),
  sub("W", "0"),
  r(" + B \u00d7 A"),  // ×
]);

module.exports = { eq1, eq2, eq3, eq4 };

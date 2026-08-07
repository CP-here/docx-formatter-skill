/**
 * mathHelpers.js
 * OMML (Office Math Markup Language) math element builders for docx library.
 *
 * Usage: const { r, sub, sup, frac, sumOp, func, paren, bracket, math, inlineMath } = require("./mathHelpers");
 *
 * CRITICAL NOTES:
 * 1. sumOp() returns an ARRAY [sumSymbol, ...body] — all functions use flat() to handle this.
 * 2. Never use MathSum with empty superScript — it renders an invisible superscript box in Word.
 *    Instead, use MathSubScript with Unicode ∑ (U+2211) as the base character.
 * 3. All children-receiving functions call flat() to prevent nested array XML corruption.
 */

const {
  Math: DMath, MathRun, MathSubScript, MathSuperScript,
  MathFraction, MathFunction, MathRoundBrackets, MathSquareBrackets,
} = require("docx");

/** Plain math text run */
function r(text) { return new MathRun(text); }

/**
 * Flatten nested arrays so sumOp's array return works everywhere.
 * Without this, nested arrays cause XML corruption and the .docx file cannot open.
 */
function flat(arr) {
  return arr.reduce((acc, val) =>
    Array.isArray(val) ? acc.concat(flat(val)) : acc.concat(val), []);
}

/** Subscript: base_{subText} e.g. sub("L", "LLM") → L_LLM */
function sub(base, subText) {
  return new MathSubScript({ children: [r(base)], subScript: [r(subText)] });
}

/** Superscript: base^{supText} e.g. sup("x", "2") → x² */
function sup(base, supText) {
  return new MathSuperScript({ children: [r(base)], superScript: [r(supText)] });
}

/** Fraction: numerator / denominator */
function frac(numChildren, denChildren) {
  return new MathFraction({ numerator: flat(numChildren), denominator: flat(denChildren) });
}

/**
 * Summation operator: ∑_{subScript} body
 *
 * IMPORTANT: Do NOT use MathSum — it requires a non-empty m:sup element,
 * which renders an empty superscript placeholder in Word.
 * Instead, use MathSubScript with Unicode ∑ (U+2211) as base character.
 *
 * Returns an array: [∑_sub, ...body]
 * The flat() function in all consumers handles this automatically.
 */
function sumOp(subScriptChildren, bodyChildren) {
  const sumSymbol = new MathSubScript({
    children: [r("\u2211")],  // Unicode ∑
    subScript: flat(subScriptChildren),
  });
  return [sumSymbol, ...flat(bodyChildren)];
}

/** Function with parentheses: funcName(args) e.g. func("log", [sub("p","i")]) → log(p_i) */
function func(name, argChildren) {
  return new MathFunction({
    name: [r(name)],
    children: [new MathRoundBrackets({ children: flat(argChildren) })],
  });
}

/** Round parentheses: (content) */
function paren(children) {
  return new MathRoundBrackets({ children: flat(children) });
}

/** Square brackets: [content] */
function bracket(children) {
  return new MathSquareBrackets({ children: flat(children) });
}

/** Block math formula (centered, standalone paragraph) */
function math(children) {
  return new DMath({ children: flat(children) });
}

/**
 * Inline math formula — can be mixed with TextRun in a Paragraph's children array.
 * Usage:
 *   new Paragraph({
 *     children: [
 *       new TextRun({ text: "其中，", font: "SimSun", size: 21 }),
 *       inlineMath([sub("L", "LLM")]),
 *       new TextRun({ text: "为语言模型损失", font: "SimSun", size: 21 }),
 *     ],
 *   });
 */
function inlineMath(children) {
  return new DMath({ children: flat(children) });
}

module.exports = { r, sub, sup, frac, sumOp, func, paren, bracket, math, inlineMath, flat };
